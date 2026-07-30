var currentHost = null;
var draggedSource = null;

var CONFIG = { teacher: "", major: "", model: "", url: "", key: "", classes: [], courses: [], defaultView: "view-home", namingVals: [], namingSeps: [], excelNamingVals: [], excelNamingSeps: [] };
var selectedClass = "";
var currentTopicCount = 3;
var generatedTopics = [];
var ITEM_HEIGHT = 43;
var uploadedImages = [];
var lastRootView = "view-home";
var lastFormData = null;
var lastAiData = null;

var AVAILABLE_VIEWS = [
    { id: "view-home", name: "主页" },
    { id: "view-main", name: "教案生成页面" },
    { id: "view-excel-main", name: "授课日志页面" }
];

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

    // 🌟 核心修复：强行隔离双平台视图，防止因为存错数据导致的卡死
    if (currentHost === Office.HostType.Word) {
        if (targetView === "view-excel-main") targetView = "view-home";
    } else if (currentHost === Office.HostType.Excel) {
        if (targetView === "view-main") targetView = "view-home";
    }

    // 🌟 如果默认是自定义页面，且已有模板数据，直接穿透到“使用”页面
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
    }

    requestAnimationFrame(function () {
        setTimeout(function () {
            injectCustomStyles();
            initStorageDeferred();
            initUIEvents();
            initDragToLoad();

            // 🌟 修复：如果设为了自定义页面，确保进入时内容能被顺利加载
            if (targetView === "view-template-use" && typeof renderTemplateUseView === 'function') {
                renderTemplateUseView();
            } else if (targetView === "view-custom" && typeof refreshCustomView === 'function') {
                refreshCustomView();
            }
        }, 10);
    });
});

function injectCustomStyles() {
    if (document.getElementById("schemaai-custom-styles")) return;
    var style = document.createElement("style");
    style.id = "schemaai-custom-styles";
    // 🌟 增加了 .topic-search-btn 的显示/隐藏逻辑
    style.innerHTML = ".image-item .img-delete { opacity: 0; background-color: #ffffff; transition: all 0.2s ease; border-radius: 50%; cursor: pointer; } .image-item:hover .img-delete { opacity: 1; } .image-item .img-delete:hover { background-color: #ff3b30 !important; } .image-item .img-delete:hover svg line { stroke: #ffffff !important; } " +
        ".topic-search-btn { display: none; } " +
        ".topic-item.selected .topic-search-btn { display: flex; align-items: center; justify-content: center; }";
    document.head.appendChild(style);
}

function updateNamingOptions() {
    // 🌟 核心修复：在数组的最前面增加一个 "无" 选项
    var options = currentHost === Office.HostType.Excel
        ? ["无", "教师授课日志", "授课专业", "授课教师", "日期"]
        : ["无", "课题", "授课教师", "日期", "教学目的"];

    document.querySelectorAll(".naming-combo-list").forEach(function (list) {
        list.innerHTML = options.map(function (opt) {
            // 选“无”的时候给个变灰样式区分
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
                // 🌟 核心修复：如果用户点击了“无”，则将输入框设为真正的空白
                input.value = text === "无" ? "" : text;
                list.classList.remove("show");
            };
        });
    });
}
function showStatus(text, type) {
    type = type || "info";
    var el = document.getElementById("status-msg");
    var el2 = document.getElementById("excel-status-msg");
    if (el) el.innerHTML = "<span class='status-" + type + "'>" + text + "</span>";
    if (el2) el2.innerHTML = "<span class='status-" + type + "'>" + text + "</span>";
}

function clearStatus() {
    var el1 = document.getElementById("status-msg");
    var el2 = document.getElementById("excel-status-msg");
    if (el1) el1.innerHTML = "";
    if (el2) el2.innerHTML = "";
}

function switchView(viewId) {
    if (viewId === "view-home" || viewId === "view-main" || viewId === "view-excel-main") lastRootView = viewId;
    document.querySelectorAll(".view-section").forEach(function (el) { el.classList.remove("active"); });
    var target = document.getElementById(viewId);
    if (target) target.classList.add("active");

    if (viewId === "view-excel-main") {
        renderExcelDailyForms();
    } else if (viewId === "view-home") {
        // 每次切回主页时，重新渲染自定义模板的专属卡片
        renderSavedTemplateCards();
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

    if (!CONFIG.timetable || CONFIG.timetable.length !== 5) {
        CONFIG.timetable = [];
        for (var d = 0; d < 5; d++) {
            var day = [];
            for (var p = 0; p < 5; p++) day.push({ course: "", classStr: "" });
            CONFIG.timetable.push(day);
        }
    }
    if (!CONFIG.expectedStudents) CONFIG.expectedStudents = {};
    if (!CONFIG.excelDailySettingsMap) CONFIG.excelDailySettingsMap = {};

    try {
        var viewSelect = document.getElementById("setting-default-view");
        if (viewSelect) {
            viewSelect.innerHTML = "";

            var allowedViews = [{ id: "view-home", name: "主页" }];
            // 🌟 核心判断：是否有可用的自定义模板数据
            var hasCustom = (CONFIG.customTemplateData && CONFIG.customTemplateData.controlsConfig && Object.keys(CONFIG.customTemplateData.controlsConfig).length > 0);

            if (currentHost === Office.HostType.Word) {
                allowedViews.push({ id: "view-main", name: "教案生成页面" });
                // 如果有，才把它的真实名字加入下拉选项
                if (hasCustom) {
                    allowedViews.push({ id: "view-template-use", name: CONFIG.customTemplateData.title || "自定义模板" });
                }
            } else if (currentHost === Office.HostType.Excel) {
                allowedViews.push({ id: "view-excel-main", name: "授课日志页面" });
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
        renderCourseDropdown();
        renderClassDropdown();
        renderTimetableUI();
        renderExpectedStudentsUI();
        renderSavedTemplateCards(); // 🌟 初始化时渲染主页专属卡片
        if (lastRootView === "view-excel-main") {
            renderExcelDailyForms();
        }
    } catch (e) { }
}

function saveStorage() { localStorage.setItem("schema_config", JSON.stringify(CONFIG)); }

function initUIEvents() {

    // Word 专用事件
    var genBtn = document.getElementById("generate-btn");
    if (genBtn) genBtn.addEventListener("click", handleGenerate);

    var saveDocBtn = document.getElementById("save-doc-btn");
    if (saveDocBtn) saveDocBtn.addEventListener("click", handleSaveDocument);

    // Excel 专用事件
    var genExcelBtn = document.getElementById("generate-excel-btn");
    if (genExcelBtn) genExcelBtn.addEventListener("click", handleExcelGenerate);

    var saveExcelBtn = document.getElementById("save-excel-btn");
    if (saveExcelBtn) saveExcelBtn.addEventListener("click", handleExcelSaveDocument);

    // 🌟 修复3：课题搜索按钮点击事件
    var searchTopicBtn = document.getElementById("search-topic-btn");
    if (searchTopicBtn) {
        searchTopicBtn.addEventListener("click", function (e) {
            e.preventDefault();
            var topicInput = document.getElementById("topic");
            var courseInput = document.getElementById("course");
            var topic = topicInput ? topicInput.value.trim() : "";
            var course = courseInput ? courseInput.value.trim() : "";

            if (!topic && !course) {
                showStatus("请先输入课题或课程", "warn");
                return;
            }
            var query = encodeURIComponent((course + " " + topic).trim() + " 教案参考");
            window.open("https://www.bing.com/search?q=" + query, "_blank");
        });
    }

    var aiTopicBtn = document.getElementById("ai-topic-btn");
    if (aiTopicBtn) {
        aiTopicBtn.addEventListener("click", function () {
            var courseInput = document.getElementById("course");
            var topicInput = document.getElementById("topic");

            clearStatus();
            if (!courseInput || !courseInput.value) {
                if (courseInput && courseInput.parentElement) courseInput.parentElement.classList.add("error-flash");
                setTimeout(function () { if (courseInput && courseInput.parentElement) courseInput.parentElement.classList.remove("error-flash"); }, 400);
                showStatus("请先填写课程名称", "error");
                return;
            }
            if (!topicInput || !topicInput.value) {
                var tw = document.getElementById("topic-wrapper");
                if (tw) tw.classList.add("error-flash");
                setTimeout(function () { if (tw) tw.classList.remove("error-flash"); }, 400);
                showStatus("请先填写章节课题大类", "error");
                return;
            }
            fetchAndRenderTopics(3, false);
        });
    }

    var topicEl = document.getElementById("topic");
    if (topicEl) {
        topicEl.addEventListener("input", function () {
            document.querySelectorAll(".topic-item").forEach(function (el) { el.classList.remove("selected"); });
        });
    }

    var navHome = document.getElementById("nav-home");
    if (navHome) navHome.addEventListener("click", function () { switchView("view-home"); });

    var navUser = document.getElementById("nav-user");
    if (navUser) {
        navUser.addEventListener("click", function () {
            var t = document.getElementById("setting-teacher");
            var m = document.getElementById("setting-major");
            if (t) t.value = CONFIG.teacher || "";
            if (m) m.value = CONFIG.major || "";
            renderTimetableUI();
            renderExpectedStudentsUI();
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

            // 回显时区分环境
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
                // 🌟 修复1：进入教案生成页面前，强制刷新下拉框数据，确保直接同步课表
                if (typeof renderCourseDropdown === 'function') renderCourseDropdown();
                if (typeof renderClassDropdown === 'function') renderClassDropdown();
                switchView("view-main");
            }
        });
    }

    var cardCustom = document.getElementById("card-custom");
    if (cardCustom) {
        cardCustom.addEventListener("click", function () {
            if (currentHost === Office.HostType.Excel) {
                var devToast = document.getElementById("dev-toast");
                if (devToast) {
                    devToast.classList.add("show");
                    setTimeout(function () { devToast.classList.remove("show"); }, 2000);
                }
                return;
            }
            // 🌟 首页固定的灰白卡片，永远只用来进入【配置】界面
            if (typeof refreshCustomView === 'function') refreshCustomView();
            switchView("view-custom");
        });
    }

    document.querySelectorAll(".nav-back-btn").forEach(function (btn) {
        btn.addEventListener("click", function (e) {
            switchView("view-home");
        });
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

            // 🌟 分离存储
            if (currentHost === Office.HostType.Excel) {
                CONFIG.excelNamingVals = tempVals;
                CONFIG.excelNamingSeps = tempSeps;
            } else {
                CONFIG.namingVals = tempVals;
                CONFIG.namingSeps = tempSeps;
            }

            saveStorage(); switchView('view-home');
        });
    }

    var saveNewCourseBtn = document.getElementById("save-new-course");
    if (saveNewCourseBtn) {
        saveNewCourseBtn.addEventListener("click", function () {
            var newCourseInput = document.getElementById("new-course-input");
            if (newCourseInput) {
                var newCourse = newCourseInput.value.trim();
                if (newCourse && (!CONFIG.courses || !CONFIG.courses.includes(newCourse))) {
                    if (!CONFIG.courses) CONFIG.courses = [];
                    CONFIG.courses.push(newCourse); saveStorage();
                }
                var ci = document.getElementById("course");
                if (ci) { ci.value = newCourse; autoSelectClassForCourse(newCourse); }
            }
            renderCourseDropdown(); switchView(lastRootView);
        });
    }

    var saveNewClassBtn = document.getElementById("save-new-class");
    if (saveNewClassBtn) {
        saveNewClassBtn.addEventListener("click", function () {
            var newClassInput = document.getElementById("new-class-input");
            if (newClassInput) {
                var newClass = newClassInput.value.trim();
                if (newClass && (!CONFIG.classes || !CONFIG.classes.includes(newClass))) {
                    if (!CONFIG.classes) CONFIG.classes = [];
                    CONFIG.classes.push(newClass); saveStorage();
                }
                var cli = document.getElementById("class-input-val");
                if (cli) cli.value = newClass;
            }
            renderClassDropdown(); switchView(lastRootView);
        });
    }

    var courseArrow = document.getElementById("course-arrow-zone");
    var courseDropdown = document.getElementById("course-select-dropdown");
    var courseDisplay = document.getElementById("course-select-display");
    if (courseArrow && courseDropdown && courseDisplay) {
        courseArrow.addEventListener("click", function (e) {
            e.stopPropagation();
            var isOpen = courseDropdown.style.display === "block";
            document.querySelectorAll(".custom-select-dropdown").forEach(function (d) { d.style.display = "none"; });
            document.querySelectorAll(".custom-select-display").forEach(function (d) { d.classList.remove("active"); });
            if (!isOpen) {
                courseDropdown.style.display = "block";
                courseDisplay.classList.add("active");
            }
        });
    }

    var classArrow = document.getElementById("class-arrow-zone");
    var classDropdown = document.getElementById("class-select-dropdown");
    var classDisplay = document.getElementById("class-select-display");
    if (classArrow && classDropdown && classDisplay) {
        classArrow.addEventListener("click", function (e) {
            e.stopPropagation();
            var isOpen = classDropdown.style.display === "block";
            document.querySelectorAll(".custom-select-dropdown").forEach(function (d) { d.style.display = "none"; });
            document.querySelectorAll(".custom-select-display").forEach(function (d) { d.classList.remove("active"); });
            if (!isOpen) {
                classDropdown.style.display = "block";
                classDisplay.classList.add("active");
            }
        });
    }

    var courseInputEl = document.getElementById("course");
    if (courseInputEl) {
        courseInputEl.addEventListener("change", function () {
            var c = this.value.trim();
            if (c) autoSelectClassForCourse(c);
        });
    }

    document.addEventListener("click", function () {
        document.querySelectorAll(".custom-select-dropdown").forEach(function (d) { d.style.display = "none"; });
        document.querySelectorAll(".custom-select-display").forEach(function (d) { d.classList.remove("active"); });
    });

    var selectDisplay = document.getElementById("class-select-display");
    var dropdown = document.getElementById("class-select-dropdown");
    if (selectDisplay && dropdown) {
        selectDisplay.addEventListener("click", function (e) {
            e.stopPropagation();
            var isOpen = dropdown.style.display === "block";
            dropdown.style.display = isOpen ? "none" : "block";
            selectDisplay.classList.toggle("active", !isOpen);
        });
        document.addEventListener("click", function () {
            dropdown.style.display = "none"; selectDisplay.classList.remove("active");
        });
    }

    var clipboardBtn = document.getElementById("clipboard-img-btn");
    if (clipboardBtn) {
        clipboardBtn.addEventListener("click", function () {
            showStatus("受限，请使用 Ctrl+V 粘贴图片", "warn");
        });
    }

    document.addEventListener("paste", function (e) {
        if (e.target.tagName === "INPUT" || e.target.tagName === "TEXTAREA") return;
        if (!e.clipboardData) return;

        var items = e.clipboardData.items;
        for (var i = 0; i < items.length; i++) {
            if (items[i].type.indexOf("image") !== -1) {
                e.preventDefault();
                var blob = items[i].getAsFile();
                var reader = new FileReader();
                reader.onload = function (event) {
                    uploadedImages.push({ id: "IMG_" + Math.random().toString(36).substr(2, 6), base64: event.target.result, desc: "" });
                    renderImageGallery();
                    showStatus("提取成功", "success");
                };
                reader.readAsDataURL(blob);
                return;
            }
        }
    });

    var addImgBtn = document.getElementById("add-img-btn");
    if (addImgBtn) {
        addImgBtn.addEventListener("click", function () {
            var input = document.getElementById("img-upload-input");
            if (input) input.click();
        });
    }

    var imgUploadInput = document.getElementById("img-upload-input");
    if (imgUploadInput) {
        imgUploadInput.addEventListener("change", function (e) {
            var files = Array.from(e.target.files);
            files.forEach(function (file) {
                var reader = new FileReader();
                reader.onload = function (event) {
                    var id = "IMG_" + Math.random().toString(36).substr(2, 6);
                    uploadedImages.push({ id: id, base64: event.target.result, desc: "" });
                    renderImageGallery();
                };
                reader.readAsDataURL(file);
            });
            e.target.value = "";
        });
    }

    var overlay = document.getElementById("modal-overlay");
    document.querySelectorAll(".btn-cancel").forEach(function (btn) { btn.addEventListener("click", function () { if (overlay) overlay.classList.remove("open"); }); });
    if (overlay) overlay.addEventListener("click", function (e) { if (e.target === overlay) overlay.classList.remove("open"); });

    var saveImgDescBtn = document.getElementById("save-img-desc");
    if (saveImgDescBtn) {
        saveImgDescBtn.addEventListener("click", function () {
            var idEl = document.getElementById("current-img-id");
            var descEl = document.getElementById("current-img-desc");
            var id = idEl ? idEl.value : "";
            var desc = descEl ? descEl.value.trim() : "";
            var img = uploadedImages.find(function (i) { return i.id === id; });
            if (img) img.desc = desc;
            if (overlay) overlay.classList.remove("open");
            renderImageGallery();
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
            } catch (err) {

            }
        });
    }
}

/* =========================================================================
   排课系统核心函数
   ========================================================================= */

function updateTimetablePlaceholders() {
    var hasAnyData = false;
    // 1. 先扫描全表，看是否有任何一个格子被填了数据
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
    // 2. 如果全表为空，只给第一个格子加提示；如果哪怕有一个字，全表提示清空
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

    for (var p = 0; p < 5; p++) {
        html += "<tr><th>第" + (p + 1) + "节</th>";
        for (var d = 0; d < 5; d++) {
            var cell = CONFIG.timetable[d][p] || { course: "", classStr: "" };
            html += "<td><div class='timetable-cell' data-day='" + d + "' data-p='" + p + "' draggable='true'><input type='text' class='tt-course' placeholder='' value='" + cell.course + "'><input type='text' class='tt-class' placeholder='' value='" + cell.classStr + "'></div></td>";
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
        cell.addEventListener("dragleave", function (e) {
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
        cell.addEventListener("dragend", function (e) {
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

function getDominantClassInfo(dayIndex) {
    var classCounts = {};
    var periods = [];

    CONFIG.timetable[dayIndex].forEach(function (cell, pIdx) {
        if (cell.classStr && cell.classStr.trim() !== "") {
            var cName = cell.classStr.trim();
            classCounts[cName] = (classCounts[cName] || 0) + 1;
            periods.push(pIdx + 1);
        }
    });

    if (Object.keys(classCounts).length === 0) return null;

    var domClass = "";
    var maxCount = -1;
    for (var cName in classCounts) {
        var count = classCounts[cName];
        if (count > maxCount) {
            maxCount = count;
            domClass = cName;
        } else if (count === maxCount) {
            var expA = parseInt(CONFIG.expectedStudents[domClass]) || 0;
            var expB = parseInt(CONFIG.expectedStudents[cName]) || 0;
            if (expB > expA) domClass = cName;
        }
    }

    var course = "";
    for (var i = 0; i < CONFIG.timetable[dayIndex].length; i++) {
        var cell = CONFIG.timetable[dayIndex][i];
        if (cell.classStr && cell.classStr.trim() === domClass && cell.course && cell.course.trim() !== "") {
            course = cell.course.trim();
            break;
        }
    }

    var expectedCount = CONFIG.expectedStudents[domClass];

    return {
        className: domClass,
        course: course,
        periodsStr: periods.join(",") + "节",
        expected: (expectedCount === "" || expectedCount === undefined) ? 0 : parseInt(expectedCount)
    };
}

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

        CONFIG.timetable[d].forEach(function (cell) {
            if (cell.course && cell.course.trim() !== "" && cell.classStr && cell.classStr.trim() !== "") {
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
    if (!CONFIG.url) throw new Error("未配置 API");

    var majorStr = CONFIG.major ? "专业：" + CONFIG.major + "。" : "";
    var prompt = "身份：专业教师。" + majorStr + "课程：" + course + "。班级：" + className + "。\n本月该班级的该课程共有 " + lessonsCount + " 次课。\n任务：生成一个涵盖全月的、具有连贯性和递进性的授课内容序列。";

    if (userInputs && userInputs.length > 0) {
        prompt += "\n参考以下用户已填写的主题进行发散、提炼和逻辑顺延：\n" + userInputs.join("；");
    }

    prompt += "\n要求：只返回纯JSON对象，格式形如 {\"contents\": [\"第1次课内容\", \"第2次课内容\", ...]}。数组长度必须恰好为 " + lessonsCount + "！内容禁止包含标点符号，严格限制在20个汉字以内。直接输出核心知识点（如：自动变速箱分类与MT构造），拒绝废话！";

    var requestBody = {
        model: CONFIG.model,
        messages: [{ role: "user", content: prompt }],
        temperature: 0.5
    };

    var isDeepSeek = /deepseek/i.test(CONFIG.model);
    var isReasoner = /reasoner|r1/i.test(CONFIG.model);
    if (isDeepSeek && !isReasoner) requestBody.response_format = { type: "json_object" };
    if (/^deepseek(-v4)?-(flash|pro)$/i.test(CONFIG.model)) requestBody.thinking = { type: "disabled" };

    var headers = { "Content-Type": "application/json" };
    if (CONFIG.key) headers["Authorization"] = "Bearer " + CONFIG.key;

    var response = await fetch(CONFIG.url, { method: "POST", headers: headers, body: JSON.stringify(requestBody) });
    if (!response.ok) throw new Error("HTTP " + response.status);

    var rawText = (await response.json()).choices[0].message.content;
    rawText = rawText.replace(/<think>[\s\S]*?<\/think>/g, "").trim().replace(/\x60\x60\x60json/gi, "").replace(/\x60\x60\x60/g, "").trim();

    var obj = JSON.parse(rawText);
    var contents = obj.contents || obj.topics || Object.values(obj)[0];
    if (Array.isArray(contents)) {
        return contents.map(function (c) { return c.replace(/[，。、！？；：“”‘’]/g, "").substring(0, 20); });
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
                if (wd >= 1 && wd <= 5) {
                    isWorkday = true;
                }
            }

            if (!isWorkday) continue;

            var dayIndex = (wd === 0 || wd === 6) ? 4 : wd - 1;

            var uniqueDayLessons = [];
            var daySeen = new Set();
            CONFIG.timetable[dayIndex].forEach(function (cell) {
                if (cell.course && cell.course.trim() !== "" && cell.classStr && cell.classStr.trim() !== "") {
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

        // 🌟 统一文案，调用刚刚声明的变量
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
                console.warn(e);
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
    if (!fileName) fileName = "授课日志";
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

/* =========================================================================
   Word 核心生成逻辑区
   ========================================================================= */
function getTimetableData() {
    var ttClasses = new Set();
    var ttCourses = new Set();
    var courseToClasses = {};

    for (var d = 0; d < 5; d++) {
        if (CONFIG.timetable && CONFIG.timetable[d]) {
            CONFIG.timetable[d].forEach(function (cell) {
                var crs = cell.course ? cell.course.trim() : "";
                var cls = cell.classStr ? cell.classStr.trim() : "";
                if (crs) ttCourses.add(crs);
                if (cls) ttClasses.add(cls);
                if (crs && cls) {
                    if (!courseToClasses[crs]) courseToClasses[crs] = new Set();
                    courseToClasses[crs].add(cls);
                }
            });
        }
    }
    return { classes: Array.from(ttClasses), courses: Array.from(ttCourses), courseMap: courseToClasses };
}

function autoSelectClassForCourse(courseName) {
    var classInput = document.getElementById("class-input-val");
    if (!classInput) return;
    var ttData = getTimetableData();
    var classesForCourse = ttData.courseMap[courseName];
    // 自动匹配，如果有多个班级则随机抽取一个填入
    if (classesForCourse && classesForCourse.size > 0) {
        var classArr = Array.from(classesForCourse);
        var randomClass = classArr[Math.floor(Math.random() * classArr.length)];
        classInput.value = randomClass;
    }
}

function renderCourseDropdown() {
    var dropdown = document.getElementById("course-select-dropdown");
    var input = document.getElementById("course");
    if (!dropdown || !input) return;

    dropdown.innerHTML = "";
    var ttData = getTimetableData();

    ttData.courses.forEach(function (c) {
        var item = document.createElement("div");
        item.className = "custom-select-item";
        item.innerHTML = "<span>" + c + " <span style='font-size:12px;color:var(--text-sub)'>(课表)</span></span>";
        item.onclick = function () {
            input.value = c;
            dropdown.style.display = "none";
            document.getElementById("course-select-display").classList.remove("active");
            autoSelectClassForCourse(c);
        };
        dropdown.appendChild(item);
    });

    if (!CONFIG.courses) CONFIG.courses = [];
    CONFIG.courses.forEach(function (c) {
        if (ttData.courses.includes(c)) return;
        var item = document.createElement("div"); item.className = "custom-select-item"; item.innerHTML = "<span>" + c + "</span>";
        var delBtn = document.createElement("span"); delBtn.className = "delete-btn";
        delBtn.innerHTML = "<svg viewBox='0 0 24 24'><line x1='18' y1='6' x2='6' y2='18'></line><line x1='6' y1='6' x2='18' y2='18'></line></svg>";
        delBtn.onclick = function (e) {
            e.stopPropagation();
            CONFIG.courses = CONFIG.courses.filter(function (crs) { return crs !== c; });
            saveStorage(); renderCourseDropdown();
        };
        item.appendChild(delBtn);
        item.onclick = function () {
            input.value = c;
            dropdown.style.display = "none";
            document.getElementById("course-select-display").classList.remove("active");
            autoSelectClassForCourse(c);
        };
        dropdown.appendChild(item);
    });

    var addItem = document.createElement("div"); addItem.className = "custom-select-item add-new"; addItem.innerText = "+ 添加课程";
    addItem.onclick = function () {
        var ni = document.getElementById("new-course-input");
        if (ni) ni.value = "";
        switchView("view-add-course");
        dropdown.style.display = "none";
        document.getElementById("course-select-display").classList.remove("active");
    };
    dropdown.appendChild(addItem);
}

function renderClassDropdown() {
    var dropdown = document.getElementById("class-select-dropdown");
    var input = document.getElementById("class-input-val");
    if (!dropdown || !input) return;

    dropdown.innerHTML = "";
    var ttData = getTimetableData();

    ttData.classes.forEach(function (c) {
        var item = document.createElement("div"); item.className = "custom-select-item";
        item.innerHTML = "<span>" + c + " <span style='font-size:12px;color:var(--text-sub)'>(课表)</span></span>";
        item.onclick = function () {
            input.value = c;
            dropdown.style.display = "none";
            document.getElementById("class-select-display").classList.remove("active");
        };
        dropdown.appendChild(item);
    });

    if (!CONFIG.classes) CONFIG.classes = [];
    CONFIG.classes.forEach(function (c) {
        if (ttData.classes.includes(c)) return;
        var item = document.createElement("div"); item.className = "custom-select-item"; item.innerHTML = "<span>" + c + "</span>";
        var delBtn = document.createElement("span"); delBtn.className = "delete-btn";
        delBtn.innerHTML = "<svg viewBox='0 0 24 24'><line x1='18' y1='6' x2='6' y2='18'></line><line x1='6' y1='6' x2='18' y2='18'></line></svg>";
        delBtn.onclick = function (e) {
            e.stopPropagation();
            CONFIG.classes = CONFIG.classes.filter(function (cls) { return cls !== c; });
            saveStorage(); renderClassDropdown();
        };
        item.appendChild(delBtn);
        item.onclick = function () {
            input.value = c;
            dropdown.style.display = "none";
            document.getElementById("class-select-display").classList.remove("active");
        };
        dropdown.appendChild(item);
    });

    var addItem = document.createElement("div"); addItem.className = "custom-select-item add-new"; addItem.innerText = "+ 添加班级";
    addItem.onclick = function () {
        var ni = document.getElementById("new-class-input");
        if (ni) ni.value = "";
        switchView("view-add-class");
        dropdown.style.display = "none";
        document.getElementById("class-select-display").classList.remove("active");
    };
    dropdown.appendChild(addItem);
}
async function fetchAndRenderTopics(targetCount, isDragging) {
    var topicInput = document.getElementById("topic");
    var courseInput = document.getElementById("course");
    var courseVal = courseInput ? courseInput.value : "";
    var panel = document.getElementById("ai-topic-panel");
    var list = document.getElementById("topic-list");
    var aiBtn = document.getElementById("ai-topic-btn");

    if (!CONFIG.url) {
        var setBtn = document.getElementById("nav-settings");
        if (setBtn) {
            setBtn.classList.add("error-flash");
            setTimeout(function () { setBtn.classList.remove("error-flash"); }, 400);
        }
        showStatus("提示：未配置 API 接口地址", "error");
        return;
    }

    currentTopicCount = targetCount;
    if (panel) panel.style.display = "block";
    if (aiBtn) aiBtn.classList.add("disabled");
    if (list) {
        list.style.height = (targetCount * ITEM_HEIGHT) + "px";
        list.innerHTML = "<div class='thinking-center'>处理中...</div>";
    }

    try {
        var prompt = "课程：" + courseVal + "。大类：" + (topicInput ? topicInput.value : "") + "。请拆分为正好 " + targetCount + " 个课时标题。难度从易到难，让学生容易理解学习。必须只返回一个纯JSON对象，格式形如 {\"topics\": [\"课题1\", \"课题2\", \"课题3\"]}。绝对禁止包含任何解释、前缀（如课题1：）、Markdown 或额外文本。";
        var headers = { "Content-Type": "application/json" };
        if (CONFIG.key) headers["Authorization"] = "Bearer " + CONFIG.key;

        var isDeepSeek = /deepseek/i.test(CONFIG.model);
        var isReasoner = /reasoner|r1/i.test(CONFIG.model);
        var isDeepSeekV4 = /^deepseek(-v4)?-(flash|pro)$/i.test(CONFIG.model);

        var requestBody = {
            model: CONFIG.model,
            messages: [{ role: "user", content: prompt }],
            temperature: 0.3
        };

        if (isDeepSeek && !isReasoner) requestBody.response_format = { type: "json_object" };
        if (isDeepSeekV4) requestBody.thinking = { type: "disabled" };

        var response = await fetch(CONFIG.url, { method: "POST", headers: headers, body: JSON.stringify(requestBody) });
        if (!response.ok) throw new Error("接口错误 (" + response.status + ")");

        var rawText = (await response.json()).choices[0].message.content;
        rawText = rawText.replace(/<think>[\s\S]*?<\/think>/g, "").trim();
        rawText = rawText.replace(/\x60{3}json|\x60{3}/gi, "").trim();

        if (rawText.startsWith("{")) {
            var obj = JSON.parse(rawText);
            var array = obj.topics || obj.items || obj.courses || Object.values(obj).find(Array.isArray);
            if (Array.isArray(array)) rawText = JSON.stringify(array);
            else throw new Error("JSON对象未找到数组");
        }

        if (!rawText.startsWith("[")) {
            var arrMatch = rawText.match(/\[([\s\S]*)\]/);
            if (arrMatch) rawText = "[" + arrMatch[1] + "]";
            else throw new Error("未找到有效数组");
        }

        generatedTopics = JSON.parse(rawText);
        if (!Array.isArray(generatedTopics)) throw new Error("结果解析错误");

        if (list) {
            list.innerHTML = "";
            generatedTopics.forEach(function (text, index) {
                var item = document.createElement("div");
                item.className = "topic-item placeholder";
                var cleanText = text.replace(/\*\*/g, "").replace(/\*/g, "");

                setTimeout(function () {
                    item.className = "topic-item slide-in selected";
                    if (index > 0) item.classList.remove("selected");
                    item.style.animationDelay = (index * 0.05) + "s";

                    // 🌟 核心修改：深绿色跳转图标（stroke='#059669'）
                    item.innerHTML = "<span class='topic-text-span' style='flex:1; pointer-events:none;'>" + cleanText + "</span>" +
                        "<button class='topic-search-btn' title='搜图片辅料' style='padding:4px; margin-left:8px; z-index:2; background:none; border:none; cursor:pointer;'><svg viewBox='0 0 24 24' width='16' height='16' stroke='#059669' stroke-width='2' fill='none' stroke-linecap='round' stroke-linejoin='round'><path d='M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6'></path><polyline points='15 3 21 3 21 9'></polyline><line x1='10' y1='14' x2='21' y2='3'></line></svg></button>";

                    var searchBtn = item.querySelector('.topic-search-btn');
                    if (searchBtn) {
                        searchBtn.onclick = function (e) {
                            e.stopPropagation();
                            // 🌟 核心修改：直接去必应搜纯粹的“图片”，关键词只有 cleanText
                            var query = encodeURIComponent(cleanText);
                            window.open("https://www.bing.com/images/search?q=" + query, "_blank");
                        };
                    }
                }, 50);

                item.onclick = function () {
                    document.querySelectorAll(".topic-item").forEach(function (el) { el.classList.remove("selected"); });
                    item.classList.add("selected");
                };
                list.appendChild(item);
            });
            setTimeout(function () { list.style.height = "auto"; }, 300);
        }
    } catch (error) {
        var errMsg = error.message || error.toString() || "未知错误";
        if (list) list.innerHTML = "<div class='thinking-center' style='color:var(--error-color);'>失败: " + errMsg + "</div>";
    } finally {
        if (aiBtn) aiBtn.classList.remove("disabled");
    }
}

function initDragToLoad() {
    var handle = document.getElementById("drag-handle");
    var list = document.getElementById("topic-list");
    if (!handle || !list) return;

    var isDragging = false, startY = 0, initialCount = 3;
    handle.addEventListener("mousedown", function (e) {
        if (generatedTopics.length === 0) return;
        isDragging = true;
        startY = e.clientY;
        initialCount = currentTopicCount;
        document.body.style.cursor = "ns-resize";
        list.style.height = (initialCount * ITEM_HEIGHT) + "px";
    });
    document.addEventListener("mousemove", function (e) {
        if (!isDragging) return;
        var deltaY = e.clientY - startY;
        var newCount = Math.max(1, initialCount + Math.round(deltaY / ITEM_HEIGHT));
        newCount = Math.min(newCount, 15);
        if (newCount !== currentTopicCount) {
            currentTopicCount = newCount;
            list.style.height = (currentTopicCount * ITEM_HEIGHT) + "px";
            var mixedData = [];
            for (var i = 0; i < currentTopicCount; i++) {
                mixedData.push(i < generatedTopics.length ? generatedTopics[i] : "");
            }
            renderDragSlots(mixedData);
        }
    });
    document.addEventListener("mouseup", function () {
        if (!isDragging) return;
        isDragging = false;
        document.body.style.cursor = "default";
        if (currentTopicCount !== generatedTopics.length) fetchAndRenderTopics(currentTopicCount, true);
        else {
            list.innerHTML = "";
            generatedTopics.forEach(function (text, index) {
                var item = document.createElement("div");
                item.className = "topic-item";
                if (index === 0) item.classList.add("selected");
                var cleanText = text.replace(/\*\*/g, "").replace(/\*/g, "");

                // 🌟 同上，只保留深绿色外链图标
                item.innerHTML = "<span class='topic-text-span' style='flex:1; pointer-events:none;'>" + cleanText + "</span>" +
                    "<button class='topic-search-btn' title='搜图片辅料' style='padding:4px; margin-left:8px; z-index:2; background:none; border:none; cursor:pointer;'><svg viewBox='0 0 24 24' width='16' height='16' stroke='#059669' stroke-width='2' fill='none' stroke-linecap='round' stroke-linejoin='round'><path d='M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6'></path><polyline points='15 3 21 3 21 9'></polyline><line x1='10' y1='14' x2='21' y2='3'></line></svg></button>";

                var searchBtn = item.querySelector('.topic-search-btn');
                if (searchBtn) {
                    searchBtn.onclick = function (e) {
                        e.stopPropagation();
                        var query = encodeURIComponent(cleanText);
                        window.open("https://www.bing.com/images/search?q=" + query, "_blank");
                    };
                }

                item.onclick = function () {
                    document.querySelectorAll(".topic-item").forEach(function (el) { el.classList.remove("selected"); });
                    item.classList.add("selected");
                };
                list.appendChild(item);
            });
            list.style.height = "auto";
        }
    });
}

function renderDragSlots(topicsArray) {
    var list = document.getElementById("topic-list");
    if (!list) return;
    list.innerHTML = "";
    topicsArray.forEach(function (text, index) {
        var item = document.createElement("div");
        item.className = "topic-item placeholder";
        if (index < generatedTopics.length) {
            item.innerText = text;
            item.style.color = "var(--text-main)";
            item.style.justifyContent = "flex-start";
        } else if (index === currentTopicCount - 1 && currentTopicCount > generatedTopics.length) {
            item.innerText = "生成 " + currentTopicCount + " 课时";
            item.style.fontWeight = "bold";
            item.style.color = "var(--text-main)";
            item.style.justifyContent = "center";
        } else {
            item.innerText = "";
        }
        list.appendChild(item);
    });
}

async function handleGenerate() {
    var btn = document.getElementById("generate-btn");
    var saveBtn = document.getElementById("save-doc-btn");
    var selectedItem = document.querySelector(".topic-item.selected");
    var topicEl = document.getElementById("topic");

    // 🌟 只读取 span 里面的纯文本，过滤掉放大镜按钮里的隐藏字符
    var selectedSpan = selectedItem ? selectedItem.querySelector(".topic-text-span") : null;
    var finalTopic = selectedSpan ? selectedSpan.innerText.trim() : (topicEl ? topicEl.value : "");

    clearStatus();
    if (saveBtn) saveBtn.classList.remove("show");

    var courseEl = document.getElementById("course");
    var classEl = document.getElementById("class-input-val");
    var currentCourse = courseEl ? courseEl.value.trim() : "";
    var currentClass = classEl ? classEl.value.trim() : "";

    if (!currentCourse) {
        if (courseEl && courseEl.parentElement) {
            courseEl.parentElement.classList.add("error-flash");
            setTimeout(function () { courseEl.parentElement.classList.remove("error-flash"); }, 400);
        }
        showStatus("请填写课程名称", "error");
        return;
    }
    if (!currentClass) {
        if (classEl && classEl.parentElement) {
            classEl.parentElement.classList.add("error-flash");
            setTimeout(function () { classEl.parentElement.classList.remove("error-flash"); }, 400);
        }
        showStatus("请选择或填写授课班级", "error");
        return;
    }
    if (!finalTopic) {
        var tw = document.getElementById("topic-wrapper");
        if (tw) {
            tw.classList.add("error-flash");
            setTimeout(function () { tw.classList.remove("error-flash"); }, 400);
        }
        showStatus("请填写或生成章节课题", "error");
        return;
    }
    if (!CONFIG.teacher || !CONFIG.major) {
        var userBtn = document.getElementById("nav-user");
        if (userBtn) {
            userBtn.classList.add("error-flash");
            setTimeout(function () { userBtn.classList.remove("error-flash"); }, 400);
        }
        showStatus("请在设置中配置教师信息", "error");
        return;
    }
    if (!CONFIG.url) {
        var setBtn = document.getElementById("nav-settings");
        if (setBtn) {
            setBtn.classList.add("error-flash");
            setTimeout(function () { setBtn.classList.remove("error-flash"); }, 400);
        }
        showStatus("请配置 API 接口", "error");
        return;
    }

    var typeRadio = document.querySelector('input[name="courseType"]:checked');
    var contentEl = document.getElementById("content");

    var formData = {
        teacher: CONFIG.teacher, major: CONFIG.major, classStr: currentClass,
        course: currentCourse, topic: finalTopic,
        courseType: typeRadio ? typeRadio.value : "理论课",
        contentInfo: contentEl ? (contentEl.value || "无") : "无",
        date: new Date().getFullYear() + "年" + String(new Date().getMonth() + 1).padStart(2, '0') + "月" + String(new Date().getDate()).padStart(2, '0') + "日"
    };
    lastFormData = formData;

    try {
        if (btn) { btn.disabled = true; btn.innerText = "处理中..."; }

        var pc = document.getElementById("progress-container");
        var pb = document.getElementById("progress-bar");
        if (pc) pc.style.display = "block";
        if (pb) pb.style.width = "10%";

        var modelNameDisplay = CONFIG.model || "AI模型";
        showStatus("正在请求 " + modelNameDisplay + "...", "info");

        let aiResponseJSON = await fetchLessonPlanFromAI(formData);

        if (formData.courseType === "理论课") {
            aiResponseJSON.practical_content = "";
            aiResponseJSON.practical_equipment = "";
        }
        lastAiData = aiResponseJSON;

        if (pb) pb.style.width = "70%";
        showStatus("等待 " + modelNameDisplay + " 响应...", "info");

        await writeToWord(formData, aiResponseJSON);

        if (pb) pb.style.width = "100%";
        showStatus("文档写入完成", "success");
        setTimeout(function () {
            if (pc) pc.style.display = "none";
            if (saveBtn) saveBtn.classList.add("show");
        }, 1000);
    } catch (error) {
        var pb2 = document.getElementById("progress-bar");
        if (pb2) pb2.style.background = "var(--error-color)";
        showStatus("运行异常：" + error.message, "error");
    } finally {
        if (btn) {
            btn.disabled = false;
            setTimeout(function () {
                btn.innerText = "执行写入";
                var pb3 = document.getElementById("progress-bar");
                if (pb3) pb3.style.background = "";
            }, 2500);
        }
    }
}

function handleSaveDocument() {
    if (!lastFormData) return;

    let validParts = [];
    for (let i = 0; i < 4; i++) {
        let val = CONFIG.namingVals[i];

        let actualStr = "";
        if (val === '课题') actualStr = lastFormData.topic || "";
        else if (val === '授课教师') actualStr = lastFormData.teacher || "";
        else if (val === '日期') actualStr = lastFormData.date || "";
        else if (val === '教学目的') actualStr = (lastAiData && lastAiData.objectives) ? lastAiData.objectives : "";
        else actualStr = val;

        if (actualStr && actualStr.trim() !== "") {
            validParts.push({ str: actualStr.trim(), originalIndex: i });
        }
    }

    let fileName = "";
    for (let j = 0; j < validParts.length; j++) {
        fileName += validParts[j].str;
        if (j < validParts.length - 1) {
            fileName += CONFIG.namingSeps[validParts[j].originalIndex];
        }
    }

    if (!fileName) fileName = "文档";
    if (lastFormData.courseType === "实训课") fileName = "实训-" + fileName;
    fileName += ".docx";

    showStatus("准备保存...", "info");
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
                            var blob = new Blob([u8Array], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });

                            (async () => {
                                try {
                                    if (window.showSaveFilePicker) {
                                        const fileHandle = await window.showSaveFilePicker({
                                            suggestedName: fileName,
                                            types: [{
                                                description: 'Word 文档',
                                                accept: { 'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx'] }
                                            }]
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

                        } else {
                            getSlice(file, ++nextSlice);
                        }
                    } else {
                        file.closeAsync();
                        showStatus("数据处理失败", "error");
                    }
                });
            }
        } else {
            showStatus("获取文档接口异常", "error");
        }
    });
}

async function fetchLessonPlanFromAI(data) {
    const jsonFormatStr = "{\"objectives\":\"目的\",\"practical_content\":\"内容\",\"practical_equipment\":\"设备\",\"focus\":\"重点\",\"difficulties\":\"难点\",\"aids\":\"辅助\",\"process_org\":\"组织\",\"process_new\":\"新课\",\"process_summary\":\"小结\",\"process_hw\":\"作业\",\"postscript\":\"后记\"}";

    let rolePrompt = "身份：上海中等职业技术学校讲师。专业" + data.major + "、课程" + data.course + "、课题" + data.topic + "。";
    if (data.courseType === "实训课") {
        rolePrompt = "身份：上海中等职业技术学校【实训指导高级教师】。正在实训场地上实训操作课。专业" + data.major + "、实操课题" + data.topic + "。";
    }

    let systemPrompt = rolePrompt + "\n请务必遵守以下规则：\n1. 仅输出单一纯JSON对象。所有的值必须是纯文本。绝对禁止包含 Markdown，不允许输出除JSON对象外其他多余语句。\n2. 教学目的：不能超过 25 字。\n3. 组织教学：写授课方式，如小游戏，举例子。\n4. 讲授新课：300-450字。每个小标题必须独占一行，以“数字. 标题内容”格式书写（如“1. 行星齿轮”），小标题后直接跟正文。禁止生成总标题和结尾字数统计！\n5. 教学后记：不超过50字的反思，禁止建议字眼。\n6. 换行规则：需要换行处输出明文 \\n，绝对禁止物理回车。\n7.  教学重点与难点要有1，2，3序号排序，每个序号独占一行，禁止用逗号分隔，每项7-15字。重点至少3个，难点至少1个。\n8.  归纳小结是对内容的重点总结，不超过50字。\n9.  教学辅助手段逗号分割。\n10. 布置简单相关作业,不要生成序号";

    if (data.courseType === "实训课") {
        systemPrompt += "\n11. 实训指令：讲授新课和组织教学必须围绕“实物拆装、设备操作”展开。不要使用太多比喻修辞，注重操作细节。绝对禁止纯理论！必须设计学生分组实操环节！";
    }

    if (data.contentInfo && data.contentInfo.trim() !== '无' && data.contentInfo.trim() !== '') {
        systemPrompt += "\n核心教学内容补充：" + data.contentInfo;
    }

    systemPrompt += "\n返回结构：" + jsonFormatStr;

    const headers = { "Content-Type": "application/json" };
    if (CONFIG.key) headers["Authorization"] = "Bearer " + CONFIG.key;

    const isDeepSeek = /deepseek/i.test(CONFIG.model);
    const isReasoner = /reasoner|r1/i.test(CONFIG.model);
    const isDeepSeekV4 = /^deepseek(-v4)?-(flash|pro)$/i.test(CONFIG.model);

    let messagesPayload = [];
    if (uploadedImages.length > 0) {
        let imgInstructions = "\n\n排版配图指令：\n前端已上传 " + uploadedImages.length + " 张图片。请仔细观察每张图片，并为需要配图的讲授新课内容，在合适位置独占一行输出对应的占位符代码：\n";
        uploadedImages.forEach((img, i) => {
            imgInstructions += (i + 1) + ". 第" + (i + 1) + "张图 (占位符：{{" + img.id + "}})\n";
        });
        imgInstructions += "请在讲授新课需要配图处，独占一行输出对应的占位符代码！";
        systemPrompt += imgInstructions;

        let contentArray = [{ type: "text", text: systemPrompt }];
        uploadedImages.forEach(img => {
            let pureBase64 = img.base64.split(',')[1];
            contentArray.push({ type: "image_url", image_url: { url: "data:image/jpeg;base64," + pureBase64 } });
        });
        messagesPayload = [{ role: "user", content: contentArray }];
    } else {
        messagesPayload = [{ role: "user", content: systemPrompt }];
    }

    async function makeRequest(payload, description = "") {
        const requestBody = {
            model: CONFIG.model,
            messages: payload,
            max_tokens: 4096
        };

        if (isDeepSeek && !isReasoner) {
            requestBody.response_format = { type: "json_object" };
        }

        if (isDeepSeekV4) {
            requestBody.thinking = { type: "enabled" };
        } else {
            requestBody.temperature = 0.7;
        }
        const response = await fetch(CONFIG.url, {
            method: "POST",
            headers: headers,
            body: JSON.stringify(requestBody)
        });
        if (!response.ok) {
            throw new Error("服务端异常 (HTTP " + response.status + ")" + (description ? ': ' + description : ''));
        }
        let rawText = (await response.json()).choices[0].message.content;
        return rawText;
    }

    function parseResponse(rawText) {
        rawText = rawText.replace(/<think>[\s\S]*?<\/think>/g, "").trim();
        rawText = rawText.replace(/```json/gi, "").replace(/```/g, "").trim();

        if (rawText.startsWith("[")) rawText = rawText.replace(/^\[/, "").replace(/\]$/, "").trim();

        let startIndex = rawText.indexOf("{");
        if (startIndex !== -1) {
            rawText = rawText.substring(startIndex);
            let openBraces = (rawText.match(/\{/g) || []).length;
            let closeBraces = (rawText.match(/\}/g) || []).length;

            while (closeBraces < openBraces) {
                if (rawText.endsWith('"')) rawText += '"}';
                else if (rawText.endsWith(',')) rawText = rawText.slice(0, -1) + '}';
                else rawText += '"}';
                closeBraces++;
            }

            let endIndex = rawText.lastIndexOf("}");
            if (endIndex !== -1) {
                rawText = rawText.substring(0, endIndex + 1);
            }
        }

        try {
            return JSON.parse(rawText);
        } catch (e) {
            let safeText = rawText.replace(/\n/g, "\\n").replace(/\r/g, "");
            safeText = safeText.replace(/,\s*\}/g, "}");
            try {
                return JSON.parse(safeText);
            } catch (e2) {
                throw new Error("模型解析异常");
            }
        }
    }

    try {
        const rawText = await makeRequest(messagesPayload);
        return parseResponse(rawText);
    } catch (error) {
        if (uploadedImages.length === 0) throw error;

        const missingDescs = uploadedImages.some(img => !img.desc || img.desc.trim() === "");
        if (missingDescs) {
            throw new Error("当前模型可能不支持多模态识别，请点击图片填写“图片简述”");
        }

        let fallbackPrompt = systemPrompt;
        let imgInstructions = "\n\n排版配图指令（无图模式）：\n已上传 " + uploadedImages.length + " 张图片，描述如下，请根据描述在讲授新课中合适的位置独占一行插入对应的占位符：\n";
        uploadedImages.forEach((img, i) => {
            imgInstructions += (i + 1) + ". 描述：\"" + (img.desc || "配图") + "\" (占位符：{{" + img.id + "}})\n";
        });
        imgInstructions += "请在讲授新课需要配图处，独占一行输出对应的占位符代码！";
        fallbackPrompt += imgInstructions;

        // 🌟 降级时也补回核心内容
        if (data.contentInfo && data.contentInfo.trim() !== '无' && data.contentInfo.trim() !== '') {
            fallbackPrompt += "\n核心教学内容补充：" + data.contentInfo;
        }

        fallbackPrompt += "\n返回结构：" + jsonFormatStr;

        const fallbackPayload = [{ role: "user", content: fallbackPrompt }];
        try {
            const rawText = await makeRequest(fallbackPayload, "纯文本降级处理");
            return parseResponse(rawText);
        } catch (retryError) {
            throw new Error("图片处理错误: " + retryError.message);
        }
    }
}

async function writeToWord(formData, aiData) {
    return Word.run(async (context) => {
        const contentMapping = {
            "cc_teacher": formData.teacher || "", "cc_date": formData.date || "", "cc_major": formData.major || "",
            "cc_class": formData.classStr || "", "cc_course": formData.course || "",
            "cc_topic": formData.topic || "", "cc_course_type": formData.courseType.replace("课", "") || "",
            "cc_objectives": String(aiData.objectives || ""), "cc_prac_content": String(aiData.practical_content || ""),
            "cc_prac_equip": String(aiData.practical_equipment || ""), "cc_focus": String(aiData.focus || ""),
            "cc_difficulties": String(aiData.difficulties || ""), "cc_aids": String(aiData.aids || ""),
            "cc_process_org": String(aiData.process_org || ""), "cc_process_new": String(aiData.process_new || ""),
            "cc_process_summary": String(aiData.process_summary || ""), "cc_process_hw": String(aiData.process_hw || ""),
            "cc_postscript": String(aiData.postscript || "")
        };

        function isTitleLine(line) {
            const trimmed = line.trim();
            return /^\d+([\.．、]\d+)*[\.．、]\s*[^\s\d]/.test(trimmed);
        }

        const allControls = context.document.contentControls;
        allControls.load("items/tag, items/title");
        await context.sync();

        let writeCount = 0;
        let imgCounter = 1;
        const targetKeys = ["cc_process_org", "cc_process_new", "cc_process_summary", "cc_process_hw", "cc_course_type"];
        const headerKeys = ["cc_teacher", "cc_date", "cc_major", "cc_class", "cc_course", "cc_topic"];

        for (const [key, rawText] of Object.entries(contentMapping)) {
            let finalRawText = rawText || " ";

            const targetControls = allControls.items.filter(c => c.tag === key || c.title === key);
            let cleanText = finalRawText.replace(/\*\*/g, "").replace(/\*/g, "").replace(/#/g, "").replace(/（字数：.*?）/g, "").replace(/\(字数[：:].*?\)/g, "");
            cleanText = cleanText.replace(/\n\s*\n+/g, "\n").trim();

            for (const targetControl of targetControls) {
                let isFirstInsert = true;

                if (key === "cc_process_new" && uploadedImages.length > 0) {
                    let parts = cleanText.split(/({{IMG_[a-zA-Z0-9]+}})/).filter(p => p !== "");

                    if (parts.length === 0) {
                        let r = targetControl.insertText(" ", "Replace");
                        r.font.spacing = 1;
                    }

                    for (let part of parts) {
                        let match = part.match(/^{{(IMG_[a-zA-Z0-9]+)}}$/);
                        if (match) {
                            let imgObj = uploadedImages.find(i => i.id === match[1]);
                            if (imgObj) {
                                let pureBase64 = imgObj.base64.split(',')[1];

                                targetControl.insertText("\n", "End");

                                let pic = targetControl.insertInlinePictureFromBase64(pureBase64, "End");
                                pic.lockAspectRatio = true;
                                pic.width = 220;

                                let descriptionStr = imgObj.desc || "辅助配图";
                                let captionText = "\n图" + imgCounter + "  " + descriptionStr;
                                let captionRange = targetControl.insertText(captionText, "End");

                                captionRange.font.name = "宋体";
                                captionRange.font.size = 10.5;
                                captionRange.font.color = "#8E8E93";
                                captionRange.font.spacing = 1;
                                captionRange.font.bold = false;

                                imgCounter++;
                                isFirstInsert = false;
                            }
                        } else {
                            let lines = part.split("\n");
                            for (let line of lines) {
                                if (line.trim() !== "") {
                                    let insertStr = isFirstInsert ? line.trim() : "\n" + line.trim();
                                    let textRange = targetControl.insertText(insertStr, isFirstInsert ? "Replace" : "End");
                                    textRange.font.name = "宋体";
                                    textRange.font.size = 12;
                                    textRange.font.color = "#1C1C1E";
                                    textRange.font.spacing = 1;
                                    textRange.font.bold = isTitleLine(line);

                                    isFirstInsert = false;
                                }
                            }
                        }
                    }
                } else {
                    let lines = cleanText.split("\n");
                    for (let line of lines) {
                        let trimmedLine = line.trim();
                        if (trimmedLine !== "" || isFirstInsert) {
                            let insertStr = isFirstInsert ? (trimmedLine || " ") : "\n" + trimmedLine;
                            let r = targetControl.insertText(insertStr, isFirstInsert ? "Replace" : "End");

                            if (headerKeys.includes(key)) {
                                r.font.name = "宋体";
                                r.font.size = 12;
                                r.font.bold = false;
                            } else if (targetKeys.includes(key)) {
                                r.font.name = "宋体";
                                r.font.size = 12;
                                r.font.color = "#1C1C1E";
                                r.font.spacing = 1;
                                r.font.bold = isTitleLine(line);
                            } else {
                                r.font.spacing = 1;
                                r.font.bold = false;
                            }
                            isFirstInsert = false;
                        }
                    }
                }
                writeCount++;
            }
        }
        await context.sync();
        if (writeCount === 0) throw new Error("未检测到占位控件");
    }).catch(e => {
        if (e instanceof OfficeExtension.Error && e.code === "AccessDenied") throw new Error("文档只读");
        throw e;
    });
}
function openImageDescModal(id) {
    var img = uploadedImages.find(function (i) { return i.id === id; });
    if (!img) return;
    document.getElementById("current-img-id").value = id;
    document.getElementById("current-img-desc").value = img.desc || "";
    var overlay = document.getElementById("modal-overlay");
    var modal = document.getElementById("img-desc-modal");
    if (overlay) overlay.classList.add("open");
    if (modal) modal.classList.add("open");
}

function renderImageGallery() {
    var gallery = document.getElementById("image-gallery");
    if (!gallery) return;
    gallery.innerHTML = "";
    uploadedImages.forEach(function (img) {
        var div = document.createElement("div");
        div.className = "image-item";
        div.onclick = function () { openImageDescModal(img.id); };
        var descHtml = img.desc ? "<div class='img-has-desc' title='已添加简述'></div>" : "";
        div.innerHTML = "<img src='" + img.base64 + "' class='img-preview' title='" + (img.desc || "点击添加图片描述") + "'>" +
            descHtml +
            "<div class='img-delete' data-id='" + img.id + "'><svg viewBox='0 0 24 24'><line x1='18' y1='6' x2='6' y2='18'></line><line x1='6' y1='6' x2='18' y2='18'></line></svg></div>";
        gallery.appendChild(div);
    });
    document.querySelectorAll(".img-delete").forEach(function (btn) {
        btn.onclick = function (e) {
            e.stopPropagation();
            var targetId = e.currentTarget.getAttribute("data-id");
            uploadedImages = uploadedImages.filter(function (i) { return i.id !== targetId; });
            renderImageGallery();
        };
    });
}
// 🌟 节假日智能抓取与缓存系统（静默处理版）
async function getHolidayConfig(year) {
    var cacheKey = "schemaai_holidays_" + year;
    var cached = localStorage.getItem(cacheKey);

    // 1. 检查本地缓存，如果有直接秒回
    if (cached) {
        try {
            return JSON.parse(cached);
        } catch (e) {
            localStorage.removeItem(cacheKey); // 缓存损坏则清除
        }
    }

    // 2. 没有缓存，静默发起网络请求
    try {
        var response = await fetch("https://api.apisbo.com/holidays/year/" + year);
        if (!response.ok) throw new Error("API 请求失败");
        var resData = await response.json();

        // 3. 构建映射表
        var holidayMap = {};

        if (resData && resData.code === 0 && Array.isArray(resData.data)) {
            resData.data.forEach(function (item) {
                var dParts = item.date.split('-');
                if (dParts.length === 3) {
                    var dStr = parseInt(dParts[0]) + "." + parseInt(dParts[1]) + "." + parseInt(dParts[2]);
                    holidayMap[dStr] = item.type;
                }
            });
        }

        // 4. 将映射表永久写入缓存
        localStorage.setItem(cacheKey, JSON.stringify(holidayMap));
        return holidayMap;

    } catch (error) {
        // 5. 只有失败的时候才会给用户提示
        console.warn("节假日 API 同步失败，降级为常规双休模式", error);
        var statusEl = document.getElementById("excel-status-msg");
        if (statusEl) statusEl.innerHTML = "<span class='status-warn'>同步节假日失败，已降级为常规工作日</span>";
        return {}; // 失败时返回空对象，走默认双休逻辑
    }
}
// 🌟 动态渲染用户已保存的专属模板卡片
function renderSavedTemplateCards() {
    var grid = document.querySelector("#view-home .template-grid");
    if (!grid) return;

    // 清除已有的动态卡片，防止重复
    var existingCards = grid.querySelectorAll(".template-card.saved-template");
    existingCards.forEach(function (c) { c.remove(); });

    var hasCustom = (CONFIG.customTemplateData && CONFIG.customTemplateData.controlsConfig && Object.keys(CONFIG.customTemplateData.controlsConfig).length > 0);

    if (hasCustom) {
        var card = document.createElement("div");
        card.className = "template-card saved-template";
        card.style.order = "-1"; // 永远排在最前面

        var title = CONFIG.customTemplateData.title || "自定义模板";
        var desc = CONFIG.customTemplateData.desc || "点击使用该模板";

        card.innerHTML =
            "<div style='position:relative;'>" +
            "<div class='card-icon' style='color: var(--success-color);'>" +
            "<svg viewBox='0 0 24 24'><rect x='3' y='3' width='18' height='18' rx='2' ry='2'></rect><line x1='3' y1='9' x2='21' y2='9'></line><line x1='9' y1='21' x2='9' y2='9'></line></svg>" +
            "</div>" +
            "<button class='delete-template-btn' title='删除模板'>" +
            "<svg viewBox='0 0 24 24' width='16' height='16' stroke='currentColor' stroke-width='2' fill='none' stroke-linecap='round'><line x1='18' y1='6' x2='6' y2='18'/><line x1='6' y1='6' x2='18' y2='18'/></svg>" +
            "</button>" +
            "</div>" +
            "<h3 style='margin: 0 0 6px 0; font-size: 13px; font-weight: 600; color: var(--text-main);'>" + title + "</h3>" +
            "<p style='margin: 0; font-size: 12px; color: var(--text-sub); line-height: 1.4;'>" + desc + "</p>";

        var delBtn = card.querySelector('.delete-template-btn');
        if (delBtn) {
            delBtn.addEventListener('click', function (e) {
                e.stopPropagation();
                // 删除数据并重置 UI
                CONFIG.customTemplateData = null;
                saveStorage();
                renderSavedTemplateCards();
                initStorageDeferred(); // 重新刷新下拉框配置
            });
        }

        // 点击卡片本体进入【使用】界面
        card.addEventListener("click", function () {
            if (typeof renderTemplateUseView === 'function') renderTemplateUseView();
            switchView("view-template-use");
        });

        grid.appendChild(card);
    }
}