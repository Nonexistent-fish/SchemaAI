// ========== Word 状态与辅助 ==========
var currentTopicCount = 3;
var generatedTopics = [];
var ITEM_HEIGHT = 43;

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
            var disp = document.getElementById("course-select-display");
            if (disp) disp.classList.remove("active");
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
            var disp = document.getElementById("course-select-display");
            if (disp) disp.classList.remove("active");
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
        var disp = document.getElementById("course-select-display");
        if (disp) disp.classList.remove("active");
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
            var disp = document.getElementById("class-select-display");
            if (disp) disp.classList.remove("active");
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
            var disp = document.getElementById("class-select-display");
            if (disp) disp.classList.remove("active");
        };
        dropdown.appendChild(item);
    });

    var addItem = document.createElement("div"); addItem.className = "custom-select-item add-new"; addItem.innerText = "+ 添加班级";
    addItem.onclick = function () {
        var ni = document.getElementById("new-class-input");
        if (ni) ni.value = "";
        switchView("view-add-class");
        dropdown.style.display = "none";
        var disp = document.getElementById("class-select-display");
        if (disp) disp.classList.remove("active");
    };
    dropdown.appendChild(addItem);
}

// ========== AI 课题拆解与外链搜图 ==========
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
        var prompt = `课程：${courseVal}。大类：${topicInput ? topicInput.value : ""}。请拆分为恰好 ${targetCount} 个循序渐进的课时标题。请仅输出纯 JSON 格式：{"topics": ["课题1", "课题2", ...]}`;
        var headers = { "Content-Type": "application/json" };
        if (CONFIG.key) headers["Authorization"] = "Bearer " + CONFIG.key;

        var requestBody = {
            model: CONFIG.model,
            messages: [{ role: "user", content: prompt }],
            temperature: 0.5
        };

        var response = await fetch(CONFIG.url, { method: "POST", headers: headers, body: JSON.stringify(requestBody) });
        var responseBodyText = await response.text();

        if (!response.ok) {
            throw new Error("接口错误 (" + response.status + "): " + responseBodyText.slice(0, 60));
        }

        var resJson;
        try {
            resJson = JSON.parse(responseBodyText);
        } catch (parseErr) {
            var cleanPreview = responseBodyText.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim().slice(0, 60);
            throw new Error("非JSON响应(请检查URL是否带/v1/chat/completions): " + (cleanPreview || "空响应"));
        }

        var rawText = (resJson.choices && resJson.choices[0] && resJson.choices[0].message)
            ? (resJson.choices[0].message.content || resJson.choices[0].message.reasoning_content || "")
            : "";

        rawText = rawText.replace(/<think>[\s\S]*?<\/think>/g, "").trim();
        rawText = rawText.replace(/\x60{3}json|\x60{3}/gi, "").trim();

        var startIndex = rawText.indexOf("{");
        var endIndex = rawText.lastIndexOf("}");
        if (startIndex !== -1 && endIndex !== -1 && endIndex > startIndex) {
            rawText = rawText.substring(startIndex, endIndex + 1);
            var obj = JSON.parse(rawText);
            generatedTopics = obj.topics || obj.items || obj.courses || Object.values(obj).find(Array.isArray) || [];
        } else {
            var arrMatch = rawText.match(/\[([\s\S]*)\]/);
            if (arrMatch) {
                generatedTopics = JSON.parse("[" + arrMatch[1] + "]");
            } else {
                throw new Error("未能解析出有效课题数组");
            }
        }

        if (!Array.isArray(generatedTopics) || generatedTopics.length === 0) throw new Error("课题列表为空");

        if (list) {
            list.innerHTML = "";
            generatedTopics.forEach(function (text, index) {
                var item = document.createElement("div");
                item.className = "topic-item placeholder";
                var cleanText = String(text).replace(/\*\*/g, "").replace(/\*/g, "").trim();

                setTimeout(function () {
                    item.className = "topic-item slide-in selected";
                    if (index > 0) item.classList.remove("selected");
                    item.style.animationDelay = (index * 0.05) + "s";

                    // 选中项右侧展示深绿色跳转图标，点击直接搜必应图片
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
        if (list) list.innerHTML = "<div class='thinking-center' style='color:var(--error-color); padding:0 12px; text-align:center;'>失败: " + errMsg + "</div>";
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
                var cleanText = String(text).replace(/\*\*/g, "").replace(/\*/g, "").trim();

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

// ========== 图片辅助交互 ==========
function openImageDescModal(id) {
    var img = uploadedImages.find(function (i) { return i.id === id; });
    if (!img) return;
    var idEl = document.getElementById("current-img-id");
    var descEl = document.getElementById("current-img-desc");
    if (idEl) idEl.value = id;
    if (descEl) descEl.value = img.desc || "";
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

// ========== Word 教案生成核心调度 ==========
async function handleGenerate() {
    var btn = document.getElementById("generate-btn");
    var saveBtn = document.getElementById("save-doc-btn");
    var selectedItem = document.querySelector(".topic-item.selected");
    var topicEl = document.getElementById("topic");

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
    var needPostscriptEl = document.getElementById("need-postscript");

    var formData = {
        teacher: CONFIG.teacher, major: CONFIG.major, classStr: currentClass,
        course: currentCourse, topic: finalTopic,
        courseType: typeRadio ? typeRadio.value : "理论课",
        contentInfo: contentEl ? (contentEl.value || "无") : "无",
        needPostscript: needPostscriptEl ? needPostscriptEl.checked : true,
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
        if (!formData.needPostscript) {
            aiResponseJSON.postscript = "";
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

async function fetchLessonPlanFromAI(data) {
    let rolePrompt = `身份：上海中等职业技术学校专业教师。专业：${data.major}，课程：${data.course}，课题：${data.topic}。`;
    if (data.courseType === "实训课") {
        rolePrompt = `身份：上海中等职业技术学校【实训指导高级教师】。正在实训车间指导实操。专业：${data.major}，实操课题：${data.topic}。`;
    }

    let systemPrompt = `${rolePrompt}
请根据上述课题进行教案设计。思考完成后，输出标准 JSON 对象，包含以下键名：
- "objectives": 教学目的（20字左右简明概述）
- "process_org": 组织教学方法（如实物演示、案例导入、互动问答）
- "process_new": 讲授新课（350字左右，按 '1. 标题\\n正文' 分段阐述）
- "focus": 教学重点（3条左右，以 1. 2. 3. 序号分行排列）
- "difficulties": 教学难点（1-2条，以 1. 2. 3. 序号分行排列）
- "aids": 教学辅助手段（逗号分隔）
- "process_summary": 归纳小结（简短总结重点）
- "process_hw": 课后作业（具体作业要求，无需序号）
- "postscript": 教学后记（教学反思）
- "practical_content": 实训内容简述（理论课可留空）
- "practical_equipment": 实训设备清单（理论课可留空）`;

    if (data.courseType === "实训课") {
        systemPrompt += "\n实训要求：讲授新课和组织教学围绕实物拆装与设备操作展开，设计学生分组实操环节。";
    }

    if (data.contentInfo && data.contentInfo.trim() !== '无' && data.contentInfo.trim() !== '') {
        systemPrompt += `\n核心教学内容补充：${data.contentInfo}`;
    }

    const headers = { "Content-Type": "application/json" };
    if (CONFIG.key) headers["Authorization"] = "Bearer " + CONFIG.key;

    let messagesPayload = [];
    if (uploadedImages.length > 0) {
        let imgInstructions = "\n\n已上传配图：\n";
        uploadedImages.forEach((img, i) => {
            imgInstructions += `${i + 1}. 配图描述: "${img.desc || '辅助配图'}" (占位符代码: {{${img.id}}})\n`;
        });
        imgInstructions += "请在 process_new（讲授新课）需要配图的位置，独占一行写入对应的占位符代码（如 {{IMG_xxx}}）。";
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

    async function makeRequest(payload) {
        const requestBody = {
            model: CONFIG.model,
            messages: payload,
            max_tokens: 16384,
            temperature: 0.6
        };

        const response = await fetch(CONFIG.url, {
            method: "POST",
            headers: headers,
            body: JSON.stringify(requestBody)
        });
        const responseText = await response.text();
        if (!response.ok) {
            throw new Error("服务端异常 (HTTP " + response.status + "): " + responseText.slice(0, 60));
        }

        let jsonRes;
        try {
            jsonRes = JSON.parse(responseText);
        } catch (err) {
            let cleanPreview = responseText.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim().slice(0, 60);
            throw new Error("非JSON响应: " + (cleanPreview || "空响应"));
        }

        let choice = jsonRes.choices && jsonRes.choices[0];
        if (!choice) throw new Error("模型未返回任何有效 choice");

        let rawContent = (choice.message && choice.message.content) ? choice.message.content : "";
        let reasoningContent = (choice.message && choice.message.reasoning_content) ? choice.message.reasoning_content : "";

        return rawContent.trim() ? rawContent : reasoningContent;
    }

    function rescueJSONFromText(text) {
        const result = {};
        const fields = ["objectives", "process_org", "process_new", "focus", "difficulties", "aids", "process_summary", "process_hw", "postscript", "practical_content", "practical_equipment"];

        fields.forEach(field => {
            const regex = new RegExp(`"${field}"\\s*:\\s*"([\\s\\S]*?)(?="\\s*,\\s*"[a-zA-Z_]+"|"\\s*\\})`, "i");
            const match = text.match(regex);
            if (match) {
                result[field] = match[1].replace(/\\n/g, "\n").replace(/\\"/g, '"').trim();
            } else {
                result[field] = "";
            }
        });

        if (result.process_new || result.objectives) {
            return result;
        }
        throw new Error("模型未能输出完整教案结构，请重新生成");
    }

    function parseResponse(rawText) {
        if (!rawText || !rawText.trim()) throw new Error("模型返回内容为空");

        rawText = rawText.replace(/<think>[\s\S]*?<\/think>/g, "").trim();

        const codeBlockMatch = rawText.match(/\x60{3}(?:json)?\s*([\s\S]*?)\s*\x60{3}/i);
        if (codeBlockMatch) {
            rawText = codeBlockMatch[1].trim();
        }

        let startIndex = rawText.indexOf("{");
        let endIndex = rawText.lastIndexOf("}");

        if (startIndex !== -1 && endIndex !== -1 && endIndex > startIndex) {
            rawText = rawText.substring(startIndex, endIndex + 1);
        }

        try {
            return JSON.parse(rawText);
        } catch (e) {
            let safeText = rawText.replace(/\r/g, "").replace(/\n/g, "\\n").replace(/,\s*\}/g, "}");
            try {
                return JSON.parse(safeText);
            } catch (e2) {
                return rescueJSONFromText(rawText);
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
            throw new Error("当前模型不支持多模态识别，请点击图片填写“辅助描述”");
        }

        const fallbackPayload = [{ role: "user", content: systemPrompt }];
        const rawText = await makeRequest(fallbackPayload);
        return parseResponse(rawText);
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
            "process_org": String(aiData.process_org || ""), "cc_process_org": String(aiData.process_org || ""),
            "cc_process_new": String(aiData.process_new || ""),
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
        const targetKeys = ["cc_process_org", "process_org", "cc_process_new", "cc_process_summary", "cc_process_hw", "cc_course_type"];
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

// ========== 页面专属事件绑定 ==========
document.addEventListener("DOMContentLoaded", function () {
    var genBtn = document.getElementById("generate-btn");
    if (genBtn) genBtn.addEventListener("click", handleGenerate);

    var saveDocBtn = document.getElementById("save-doc-btn");
    if (saveDocBtn) saveDocBtn.addEventListener("click", handleSaveDocument);

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
});