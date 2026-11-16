// ========== Office 启动与环境初始化 ==========
Office.onReady(function (info) {
    currentHost = info.host;

    var badge = document.getElementById("host-badge");
    if (badge) {
        if (info.host === Office.HostType.Word) {
            badge.textContent = "Word";
            badge.className = "host-badge word";
        } else if (info.host === Office.HostType.Excel) {
            badge.textContent = "Excel";
            badge.className = "host-badge excel";
        }
    }

    var targetView = "view-home";
    try {
        var stored = localStorage.getItem("schema_config");
        if (stored) {
            var parsed = JSON.parse(stored);
            if (parsed.defaultView) targetView = parsed.defaultView;
        }
    } catch (e) { }

    // 隔离双平台主视图
    if (currentHost === Office.HostType.Word) {
        if (targetView === "view-excel-main" || targetView === "view-excel-plan") targetView = "view-home";
    } else if (currentHost === Office.HostType.Excel) {
        if (targetView === "view-main") targetView = "view-home";
    }

    if (targetView === "view-custom") {
        try {
            var storedData = localStorage.getItem("schema_config");
            if (storedData) {
                var pData = JSON.parse(storedData);
                if (pData.customTemplateData && pData.customTemplateData.controlsConfig && Object.keys(pData.customTemplateData.controlsConfig).length > 0) {
                    targetView = "view-template-use";
                }
            }
        } catch (e) { }
    }

    var targetEl = document.getElementById(targetView);
    if (!targetEl) targetView = "view-home";

    document.querySelectorAll(".view-section").forEach(function (el) { el.classList.remove("active"); });
    document.getElementById(targetView).classList.add("active");
    lastRootView = targetView;

    if (currentHost === Office.HostType.Excel) {
        var titleEl = document.getElementById("primary-tool-title");
        var descEl = document.getElementById("primary-tool-desc");
        if (titleEl) titleEl.innerText = "授课日志";
        if (descEl) descEl.innerText = "自动生成授课日志";

        // 在 Excel 模式下展示「授课计划」主页卡片
        var planCard = document.getElementById("card-excel-plan");
        if (planCard) planCard.style.display = "block";
    }

    requestAnimationFrame(function () {
        setTimeout(function () {
            injectCustomStyles();
            initStorageDeferred();
            initUIEvents();
            if (typeof initDragToLoad === 'function') initDragToLoad();

            if (targetView === "view-template-use" && typeof renderTemplateUseView === 'function') {
                renderTemplateUseView();
            } else if (targetView === "view-custom" && typeof refreshCustomView === 'function') {
                refreshCustomView();
            } else if (targetView === "view-excel-plan" && typeof renderExcelPlanView === 'function') {
                renderExcelPlanView();
            }
        }, 10);
    });
});

function injectCustomStyles() {
    if (document.getElementById("schemaai-custom-styles")) return;
    var style = document.createElement("style");
    style.id = "schemaai-custom-styles";
    style.innerHTML = ".image-item .img-delete { opacity: 0; background-color: #ffffff; transition: all 0.2s ease; border-radius: 50%; cursor: pointer; } .image-item:hover .img-delete { opacity: 1; } .image-item .img-delete:hover { background-color: #ff3b30 !important; } .image-item .img-delete:hover svg line { stroke: #ffffff !important; } " +
        ".topic-search-btn { display: none; } " +
        ".topic-item.selected .topic-search-btn { display: flex; align-items: center; justify-content: center; }";
    document.head.appendChild(style);
}

function switchView(viewId) {
    if (viewId === "view-home" || viewId === "view-main" || viewId === "view-excel-main" || viewId === "view-excel-plan") {
        lastRootView = viewId;
    }
    document.querySelectorAll(".view-section").forEach(function (el) { el.classList.remove("active"); });
    var target = document.getElementById(viewId);
    if (target) target.classList.add("active");

    if (viewId === "view-excel-main") {
        if (typeof renderExcelDailyForms === 'function') renderExcelDailyForms();
    } else if (viewId === "view-excel-plan") {
        if (typeof renderExcelPlanView === 'function') renderExcelPlanView();
    } else if (viewId === "view-home") {
        if (typeof renderSavedTemplateCards === 'function') renderSavedTemplateCards();
    }
}

function initStorageDeferred() {
    try {
        var stored = localStorage.getItem("schema_config");
        if (stored) {
            var parsed = JSON.parse(stored);
            CONFIG = Object.assign({}, CONFIG, parsed);
            if (Array.isArray(CONFIG.classes)) {
                CONFIG.classes = CONFIG.classes.filter(function (c) { return c && String(c).trim() !== ""; });
            } else { CONFIG.classes = []; }

            if (Array.isArray(CONFIG.courses)) {
                CONFIG.courses = CONFIG.courses.filter(function (c) { return c && String(c).trim() !== ""; });
            } else { CONFIG.courses = []; }
        }
    } catch (e) { }

    if (!CONFIG.namingVals || CONFIG.namingVals.length !== 4) {
        CONFIG.namingVals = ["课题", "", "日期", ""];
        CONFIG.namingSeps = ["-", "-", "-"];
    }
    if (!CONFIG.excelNamingVals || CONFIG.excelNamingVals.length !== 4) {
        CONFIG.excelNamingVals = ["教师授课日志", "", "日期", ""];
        CONFIG.excelNamingSeps = ["-", "-", "-"];
    }

    // 🌟 课程表升级为 6行×5列（周一至周五 × 每天 6 节课），并兼容旧版 5 节课存档自动扩容
    if (!CONFIG.timetable || CONFIG.timetable.length !== 5) {
        CONFIG.timetable = [];
        for (var d = 0; d < 5; d++) {
            var day = [];
            for (var p = 0; p < 6; p++) day.push({ course: "", classStr: "" });
            CONFIG.timetable.push(day);
        }
    } else {
        for (var dIdx = 0; dIdx < 5; dIdx++) {
            if (!Array.isArray(CONFIG.timetable[dIdx])) CONFIG.timetable[dIdx] = [];
            while (CONFIG.timetable[dIdx].length < 6) {
                CONFIG.timetable[dIdx].push({ course: "", classStr: "" });
            }
        }
    }

    if (!CONFIG.expectedStudents) CONFIG.expectedStudents = {};
    if (!CONFIG.excelDailySettingsMap) CONFIG.excelDailySettingsMap = {};
    if (!Array.isArray(CONFIG.customSyllabuses)) CONFIG.customSyllabuses = [];
    if (!CONFIG.excelPlanMap) CONFIG.excelPlanMap = {};

    try {
        var viewSelect = document.getElementById("setting-default-view");
        if (viewSelect) {
            viewSelect.innerHTML = "";
            var allowedViews = [{ id: "view-home", name: "主页" }];
            var hasCustom = (CONFIG.customTemplateData && CONFIG.customTemplateData.controlsConfig && Object.keys(CONFIG.customTemplateData.controlsConfig).length > 0);

            if (currentHost === Office.HostType.Word) {
                allowedViews.push({ id: "view-main", name: "教案生成页面" });
                if (hasCustom) {
                    allowedViews.push({ id: "view-template-use", name: CONFIG.customTemplateData.title || "自定义模板" });
                }
            } else if (currentHost === Office.HostType.Excel) {
                allowedViews.push({ id: "view-excel-main", name: "授课日志页面" });
                allowedViews.push({ id: "view-excel-plan", name: "授课计划页面" });
                if (hasCustom) {
                    allowedViews.push({ id: "view-template-use", name: CONFIG.customTemplateData.title || "自定义模板" });
                }
            }

            allowedViews.forEach(function (v) {
                var opt = document.createElement("option");
                opt.value = v.id;
                opt.innerText = v.name;
                viewSelect.appendChild(opt);
            });

            var isValid = allowedViews.find(function (v) { return v.id === CONFIG.defaultView; });
            CONFIG.defaultView = isValid ? CONFIG.defaultView : "view-home";
            viewSelect.value = CONFIG.defaultView;
        }
    } catch (e) { }

    updateNamingOptions();

    try {
        var isExcel = currentHost === Office.HostType.Excel;
        var currentVals = isExcel ? CONFIG.excelNamingVals : CONFIG.namingVals;
        var currentSeps = isExcel ? CONFIG.excelNamingSeps : CONFIG.namingSeps;

        for (var i = 0; i < 4; i++) {
            var input = document.getElementById("naming-val-" + i);
            if (input) input.value = currentVals[i] || "";
            if (i < 3) {
                var sepBtn = document.getElementById("naming-sep-" + i);
                if (sepBtn) {
                    var val = currentSeps[i] || "-";
                    sepBtn.dataset.val = val;
                    sepBtn.innerText = val === " " ? "空格" : val;
                }
            }
        }
        if (typeof renderCourseDropdown === 'function') renderCourseDropdown();
        if (typeof renderClassDropdown === 'function') renderClassDropdown();
        if (typeof renderTimetableUI === 'function') renderTimetableUI();
        if (typeof renderExpectedStudentsUI === 'function') renderExpectedStudentsUI();
        if (typeof renderSavedTemplateCards === 'function') renderSavedTemplateCards();
        if (lastRootView === "view-excel-main" && typeof renderExcelDailyForms === 'function') {
            renderExcelDailyForms();
        } else if (lastRootView === "view-excel-plan" && typeof renderExcelPlanView === 'function') {
            renderExcelPlanView();
        }
    } catch (e) { }
}

function updateNamingOptions() {
    var options = currentHost === Office.HostType.Excel
        ? ["无", "教师授课日志", "授课计划进度表", "授课专业", "授课教师", "日期"]
        : ["无", "课题", "授课教师", "日期", "教学目的"];

    document.querySelectorAll(".naming-combo-list").forEach(function (list) {
        list.innerHTML = options.map(function (opt) {
            var styleStr = opt === "无" ? "style='color:var(--text-sub); font-style:italic;'" : "";
            return "<div class='naming-combo-item' " + styleStr + ">" + opt + "</div>";
        }).join("");
    });
    bindNamingComboEvents();
}

function bindNamingComboEvents() {
    document.querySelectorAll(".naming-combo").forEach(function (combo) {
        var input = combo.querySelector("input");
        var list = combo.querySelector(".naming-combo-list");
        if (!input || !list) return;

        input.onclick = function (e) {
            e.stopPropagation();
            document.querySelectorAll(".naming-combo-list").forEach(function (l) { if (l !== list) l.classList.remove("show"); });
            list.classList.toggle("show");
        };

        list.querySelectorAll(".naming-combo-item").forEach(function (item) {
            item.onclick = function (e) {
                e.stopPropagation();
                var text = item.innerText;
                input.value = text === "无" ? "" : text;
                list.classList.remove("show");
            };
        });
    });
}

function initUIEvents() {
    var navHome = document.getElementById("nav-home");
    if (navHome) navHome.addEventListener("click", function () { switchView("view-home"); });

    var navUser = document.getElementById("nav-user");
    if (navUser) {
        navUser.addEventListener("click", function () {
            var t = document.getElementById("setting-teacher");
            var m = document.getElementById("setting-major");
            if (t) t.value = CONFIG.teacher || "";
            if (m) m.value = CONFIG.major || "";
            if (typeof renderTimetableUI === 'function') renderTimetableUI();
            if (typeof renderExpectedStudentsUI === 'function') renderExpectedStudentsUI();
            switchView("view-user");
        });
    }

    var navSettings = document.getElementById("nav-settings");
    if (navSettings) {
        navSettings.addEventListener("click", function () {
            var dv = document.getElementById("setting-default-view");
            var sm = document.getElementById("setting-model");
            var su = document.getElementById("setting-url");
            var sk = document.getElementById("setting-key");
            if (dv) dv.value = CONFIG.defaultView || "view-home";
            if (sm) sm.value = CONFIG.model || "";
            if (su) su.value = CONFIG.url || "";
            if (sk) sk.value = CONFIG.key || "";

            var isExcel = currentHost === Office.HostType.Excel;
            var currentVals = isExcel ? CONFIG.excelNamingVals : CONFIG.namingVals;
            var currentSeps = isExcel ? CONFIG.excelNamingSeps : CONFIG.namingSeps;

            for (var i = 0; i < 4; i++) {
                var input = document.getElementById("naming-val-" + i);
                if (input) input.value = currentVals[i] || "";
                if (i < 3) {
                    var sepBtn = document.getElementById("naming-sep-" + i);
                    if (sepBtn) {
                        var val = currentSeps[i] || "-";
                        sepBtn.dataset.val = val;
                        sepBtn.innerText = val === " " ? "空格" : val;
                    }
                }
            }
            switchView("view-settings");
        });
    }

    var cardPrimary = document.getElementById("card-primary-tool");
    if (cardPrimary) {
        cardPrimary.addEventListener("click", function () {
            if (currentHost === Office.HostType.Excel) {
                switchView("view-excel-main");
            } else {
                if (typeof renderCourseDropdown === 'function') renderCourseDropdown();
                if (typeof renderClassDropdown === 'function') renderClassDropdown();
                switchView("view-main");
            }
        });
    }

    // Excel 授课计划卡片点击进入
    var cardExcelPlan = document.getElementById("card-excel-plan");
    if (cardExcelPlan) {
        cardExcelPlan.addEventListener("click", function () {
            switchView("view-excel-plan");
        });
    }

    var cardCustom = document.getElementById("card-custom");
    if (cardCustom) {
        cardCustom.addEventListener("click", function () {
            if (typeof refreshCustomView === 'function') refreshCustomView();
            switchView("view-custom");
        });
    }

    document.querySelectorAll(".nav-back-btn").forEach(function (btn) {
        btn.addEventListener("click", function () { switchView("view-home"); });
    });

    var saveUserBtn = document.getElementById("save-user");
    if (saveUserBtn) {
        saveUserBtn.addEventListener("click", function () {
            var t = document.getElementById("setting-teacher");
            var m = document.getElementById("setting-major");
            if (t) CONFIG.teacher = t.value.trim();
            if (m) CONFIG.major = m.value.trim();
            saveStorage();
            showStatus("配置保存成功", "success");
            setTimeout(function () { switchView("view-home"); }, 500);
        });
    }

    document.addEventListener("click", function () {
        document.querySelectorAll(".naming-combo-list").forEach(function (l) { l.classList.remove("show"); });
    });

    var sepTypes = ["-", "+", " ", ","];
    var sepLabels = ["-", "+", "空格", ","];
    for (var i = 0; i < 3; i++) {
        var sepBtn = document.getElementById("naming-sep-" + i);
        if (sepBtn) {
            sepBtn.addEventListener("click", function (e) {
                var btn = e.currentTarget;
                var current = btn.dataset.val || "-";
                var idx = sepTypes.indexOf(current);
                var nextIdx = (idx + 1) % sepTypes.length;
                btn.dataset.val = sepTypes[nextIdx];
                btn.innerText = sepLabels[nextIdx];
            });
        }
    }

    var saveSettingsBtn = document.getElementById("save-settings");
    if (saveSettingsBtn) {
        saveSettingsBtn.addEventListener("click", function () {
            var dv = document.getElementById("setting-default-view");
            var sm = document.getElementById("setting-model");
            var su = document.getElementById("setting-url");
            var sk = document.getElementById("setting-key");
            if (dv) CONFIG.defaultView = dv.value;
            if (sm) CONFIG.model = sm.value.trim();
            if (su) CONFIG.url = su.value.trim();
            if (sk) CONFIG.key = sk.value.trim();

            var tempVals = [];
            var tempSeps = [];
            for (var j = 0; j < 4; j++) {
                var v = document.getElementById("naming-val-" + j);
                if (v) tempVals.push(v.value);
                if (j < 3) {
                    var s = document.getElementById("naming-sep-" + j);
                    if (s) tempSeps.push(s.dataset.val);
                }
            }

            if (currentHost === Office.HostType.Excel) {
                CONFIG.excelNamingVals = tempVals;
                CONFIG.excelNamingSeps = tempSeps;
            } else {
                CONFIG.namingVals = tempVals;
                CONFIG.namingSeps = tempSeps;
            }

            saveStorage();
            switchView('view-home');
        });
    }

    var browsePathBtn = document.getElementById("browse-path-btn");
    if (browsePathBtn) {
        browsePathBtn.addEventListener("click", async function () {
            try {
                if (window.showDirectoryPicker) {
                    var dirHandle = await window.showDirectoryPicker();
                    var pathInput = document.getElementById("setting-save-path");
                    if (pathInput) pathInput.value = dirHandle.name;
                    showStatus("已选取目录: " + dirHandle.name + " (受限无法获取绝对路径)", "success");
                } else {
                    showStatus("当前环境受限，请手动复制绝对路径粘贴", "warn");
                }
            } catch (err) { }
        });
    }
}