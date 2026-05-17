Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        initStorage();
        initUIEvents();
        initDragToLoad();
    }
});

let CONFIG = { teacher: "", major: "", model: "", url: "", key: "", classes: [], defaultView: "view-home" };
let selectedClass = "";
let currentTopicCount = 3;
let generatedTopics = [];
const ITEM_HEIGHT = 43;
let uploadedImages = [];

let lastRootView = "view-home"; // 记录跳转前的根视图（主页或教案）

// 可用启动页面列表
const AVAILABLE_VIEWS = [
    { id: "view-home", name: "主页" },
    { id: "view-main", name: "教案生成页面" }
];

function showStatus(text, type = "info") { document.getElementById("status-msg").innerHTML = `<span class="status-${type}">${text}</span>`; }
function clearStatus() { document.getElementById("status-msg").innerHTML = ""; }

function switchView(viewId) {

    if (viewId === "view-home" || viewId === "view-main") {
        lastRootView = viewId;
    }

    document.querySelectorAll(".view-section").forEach(el => el.classList.remove("active"));
    document.getElementById(viewId).classList.add("active");
}

function initStorage() {
    const stored = localStorage.getItem("schema_config");
    if (stored) CONFIG = { ...CONFIG, ...JSON.parse(stored) };
    else {
        CONFIG.model = "";
        CONFIG.url = "";
        CONFIG.classes = [""];
        CONFIG.defaultView = "view-home";
    }

    const viewSelect = document.getElementById("setting-default-view");
    viewSelect.innerHTML = "";
    AVAILABLE_VIEWS.forEach(v => {
        let opt = document.createElement("option");
        opt.value = v.id;
        opt.innerText = v.name;
        viewSelect.appendChild(opt);
    });

    if (AVAILABLE_VIEWS.find(v => v.id === CONFIG.defaultView)) {
        switchView(CONFIG.defaultView);
    } else {
        switchView("view-home");
    }

    renderClassDropdown();
}

function saveStorage() { localStorage.setItem("schema_config", JSON.stringify(CONFIG)); }

function showDevToast() {
    const toast = document.getElementById("dev-toast");
    toast.classList.add("show");
    setTimeout(() => toast.classList.remove("show"), 2000);
}

function initUIEvents() {
    document.getElementById("generate-btn").onclick = handleGenerate;
    document.getElementById("ai-topic-btn").onclick = () => fetchAndRenderTopics(3, false);

    // 顶部导航触发
    document.getElementById("nav-home").onclick = () => switchView("view-home");
    document.getElementById("nav-user").onclick = () => {
        document.getElementById("setting-teacher").value = CONFIG.teacher;
        document.getElementById("setting-major").value = CONFIG.major;
        switchView("view-user");
    };
    document.getElementById("nav-settings").onclick = () => {
        document.getElementById("setting-default-view").value = CONFIG.defaultView;
        document.getElementById("setting-model").value = CONFIG.model;
        document.getElementById("setting-url").value = CONFIG.url;
        document.getElementById("setting-key").value = CONFIG.key;
        switchView("view-settings");
    };

    document.getElementById("card-lesson-plan").onclick = () => switchView("view-main");
    document.getElementById("card-custom").onclick = showDevToast;

    document.querySelectorAll(".nav-back-btn").forEach(btn => {
        btn.onclick = () => switchView(lastRootView);
    });

    document.getElementById("save-user").onclick = () => {
        CONFIG.teacher = document.getElementById("setting-teacher").value.trim();
        CONFIG.major = document.getElementById("setting-major").value.trim();
        saveStorage(); switchView(lastRootView);
    };
    document.getElementById("save-settings").onclick = () => {
        CONFIG.defaultView = document.getElementById("setting-default-view").value;
        CONFIG.model = document.getElementById("setting-model").value.trim();
        CONFIG.url = document.getElementById("setting-url").value.trim();
        CONFIG.key = document.getElementById("setting-key").value.trim();
        saveStorage(); switchView(lastRootView);
    };
    document.getElementById("save-new-class").onclick = () => {
        const newClass = document.getElementById("new-class-input").value.trim();
        if (newClass && !CONFIG.classes.includes(newClass)) {
            CONFIG.classes.push(newClass); saveStorage(); selectedClass = newClass;
        }
        renderClassDropdown(); switchView(lastRootView);
    };

    const selectDisplay = document.getElementById("class-select-display");
    const dropdown = document.getElementById("class-select-dropdown");
    selectDisplay.onclick = (e) => {
        e.stopPropagation();
        const isOpen = dropdown.style.display === "block";
        dropdown.style.display = isOpen ? "none" : "block";
        selectDisplay.classList.toggle("active", !isOpen);
    };
    document.addEventListener("click", () => {
        dropdown.style.display = "none"; selectDisplay.classList.remove("active");
    });

    document.getElementById("add-img-btn").onclick = () => document.getElementById("img-upload-input").click();
    document.getElementById("img-upload-input").onchange = (e) => {
        const files = Array.from(e.target.files);
        files.forEach(file => {
            const reader = new FileReader();
            reader.onload = (event) => {
                const id = "IMG_" + Math.random().toString(36).substr(2, 6);
                uploadedImages.push({ id: id, base64: event.target.result, desc: "" });
                renderImageGallery();
            };
            reader.readAsDataURL(file);
        });
        e.target.value = "";
    };

    const overlay = document.getElementById("modal-overlay");
    document.querySelectorAll(".btn-cancel").forEach(btn => btn.onclick = () => overlay.classList.remove("open"));
    overlay.onclick = (e) => { if (e.target === overlay) overlay.classList.remove("open"); };

    document.getElementById("save-img-desc").onclick = () => {
        const id = document.getElementById("current-img-id").value;
        const desc = document.getElementById("current-img-desc").value.trim();
        const img = uploadedImages.find(i => i.id === id);
        if (img) img.desc = desc;
        overlay.classList.remove("open");
        renderImageGallery();
    };
}

function openImageDescModal(id) {
    const img = uploadedImages.find(i => i.id === id);
    if (!img) return;
    document.getElementById("current-img-id").value = id;
    document.getElementById("current-img-desc").value = img.desc || "";
    document.getElementById("modal-overlay").classList.add("open");
    document.getElementById("img-desc-modal").classList.add("open");
}

function renderImageGallery() {
    const gallery = document.getElementById("image-gallery");
    gallery.innerHTML = "";
    uploadedImages.forEach(img => {
        const div = document.createElement("div"); div.className = "image-item";
        div.onclick = () => openImageDescModal(img.id);
        div.innerHTML = `
            <img src="${img.base64}" class="img-preview" title="${img.desc || '点击添加图片描述'}">
            ${img.desc ? '<div class="img-has-desc" title="已添加简述"></div>' : ''}
            <div class="img-delete" data-id="${img.id}"><svg viewBox="0 0 24 24"><line x1="18" y1="6" x2="6" y2="18"></line><line x1="6" y1="6" x2="18" y2="18"></line></svg></div>
        `;
        gallery.appendChild(div);
    });
    document.querySelectorAll(".img-delete").forEach(btn => {
        btn.onclick = (e) => {
            e.stopPropagation();
            uploadedImages = uploadedImages.filter(i => i.id !== e.currentTarget.getAttribute("data-id"));
            renderImageGallery();
        };
    });
}

function renderClassDropdown() {
    const dropdown = document.getElementById("class-select-dropdown");
    const display = document.getElementById("class-select-display-text");
    dropdown.innerHTML = "";
    if (CONFIG.classes.length === 0) selectedClass = "";
    else if (!selectedClass || !CONFIG.classes.includes(selectedClass)) selectedClass = CONFIG.classes[0];
    display.innerText = selectedClass || "请添加班级";

    CONFIG.classes.forEach(c => {
        const item = document.createElement("div"); item.className = "custom-select-item"; item.innerHTML = `<span>${c}</span>`;
        const delBtn = document.createElement("span"); delBtn.className = "delete-btn";
        delBtn.innerHTML = `<svg viewBox="0 0 24 24"><line x1="18" y1="6" x2="6" y2="18"></line><line x1="6" y1="6" x2="18" y2="18"></line></svg>`;
        delBtn.onclick = (e) => {
            e.stopPropagation(); CONFIG.classes = CONFIG.classes.filter(cls => cls !== c);
            if (selectedClass === c) selectedClass = ""; saveStorage(); renderClassDropdown();
        };
        item.appendChild(delBtn);
        item.onclick = () => { selectedClass = c; display.innerText = c; };
        dropdown.appendChild(item);
    });
    const addItem = document.createElement("div"); addItem.className = "custom-select-item add-new"; addItem.innerText = "+ 添加班级";
    addItem.onclick = () => { document.getElementById("new-class-input").value = ""; switchView("view-add-class"); };
    dropdown.appendChild(addItem);
}

// AI 课题拆解
async function fetchAndRenderTopics(targetCount, isDragging) {
    const topicInput = document.getElementById("topic");
    const courseInput = document.getElementById("course").value;
    const panel = document.getElementById("ai-topic-panel");
    const list = document.getElementById("topic-list");
    const aiBtn = document.getElementById("ai-topic-btn");

    if (!CONFIG.url) { alert(" 错误：未配置 API 接口地址"); return; }
    if (!topicInput.value || !courseInput) {
        document.getElementById("topic-wrapper").classList.add("error-flash");
        setTimeout(() => document.getElementById("topic-wrapper").classList.remove("error-flash"), 400); return;
    }
    currentTopicCount = targetCount; panel.style.display = "block"; aiBtn.classList.add("disabled");
    list.style.height = `${targetCount * ITEM_HEIGHT}px`; list.innerHTML = `<div class="thinking-center">Thinking...</div>`;

    try {
        const prompt = `课程：${courseInput}。大类：${topicInput.value}。请拆分为正好 ${targetCount} 个课时标题。难度从易到难，让学生容易理解学习，只返回纯JSON数组，绝对禁止 Markdown。`;
        const headers = { "Content-Type": "application/json" };
        if (CONFIG.key) headers["Authorization"] = `Bearer ${CONFIG.key}`;

        const response = await fetch(CONFIG.url, { method: "POST", headers: headers, body: JSON.stringify({ model: CONFIG.model, messages: [{ role: "user", content: prompt }], temperature: 0.3 }) });
        if (!response.ok) throw new Error("接口连接被拒");
        let rawText = (await response.json()).choices[0].message.content.replace(/```json/g, '').replace(/```/g, '').trim();
        generatedTopics = JSON.parse(rawText);
        list.innerHTML = "";
        generatedTopics.forEach((text, index) => {
            const item = document.createElement("div"); item.className = "topic-item placeholder";
            setTimeout(() => {
                item.className = "topic-item slide-in selected"; if (index > 0) item.classList.remove("selected");
                item.style.animationDelay = `${index * 0.05}s`; item.innerText = text.replace(/\*\*/g, "").replace(/\*/g, "");
            }, 50);
            item.onclick = () => { document.querySelectorAll(".topic-item").forEach(el => el.classList.remove("selected")); item.classList.add("selected"); };
            list.appendChild(item);
        });
        setTimeout(() => { list.style.height = "auto"; }, 300);
    } catch (error) { list.innerHTML = `<div class="thinking-center" style="color:var(--error-color);">失败: ${error.message}</div>`; } finally { aiBtn.classList.remove("disabled"); }
}

function initDragToLoad() {
    const handle = document.getElementById("drag-handle"); const list = document.getElementById("topic-list");
    let isDragging = false, startY = 0, initialCount = 3;
    handle.addEventListener("mousedown", (e) => {
        if (generatedTopics.length === 0) return; isDragging = true; startY = e.clientY; initialCount = currentTopicCount; document.body.style.cursor = "ns-resize"; list.style.height = `${initialCount * ITEM_HEIGHT}px`;
    });
    document.addEventListener("mousemove", (e) => {
        if (!isDragging) return;
        const deltaY = e.clientY - startY; let newCount = Math.max(1, initialCount + Math.round(deltaY / ITEM_HEIGHT)); newCount = Math.min(newCount, 15);
        if (newCount !== currentTopicCount) { currentTopicCount = newCount; list.style.height = `${currentTopicCount * ITEM_HEIGHT}px`; const mixedData = []; for (let i = 0; i < currentTopicCount; i++) { mixedData.push(i < generatedTopics.length ? generatedTopics[i] : ""); } renderDragSlots(mixedData); }
    });
    document.addEventListener("mouseup", () => {
        if (!isDragging) return; isDragging = false; document.body.style.cursor = "default";
        if (currentTopicCount !== generatedTopics.length) fetchAndRenderTopics(currentTopicCount, true);
        else {
            list.innerHTML = ""; generatedTopics.forEach((text, index) => {
                const item = document.createElement("div"); item.className = "topic-item"; if (index === 0) item.classList.add("selected"); item.innerText = text;
                item.onclick = () => { document.querySelectorAll(".topic-item").forEach(el => el.classList.remove("selected")); item.classList.add("selected"); };
                list.appendChild(item);
            });
            list.style.height = "auto";
        }
    });
}
function renderDragSlots(topicsArray) {
    const list = document.getElementById("topic-list"); list.innerHTML = "";
    topicsArray.forEach((text, index) => {
        const item = document.createElement("div"); item.className = "topic-item placeholder";
        if (index < generatedTopics.length) { item.innerText = text; item.style.color = "var(--text-main)"; item.style.justifyContent = "flex-start"; }
        else if (index === currentTopicCount - 1 && currentTopicCount > generatedTopics.length) { item.innerText = `分类生成 ${currentTopicCount} 课时`; item.style.fontWeight = "bold"; item.style.color = "var(--text-main)"; item.style.justifyContent = "center"; }
        else { item.innerText = ""; }
        list.appendChild(item);
    });
}

let progressInterval = null;
function updateProgress(width) { document.getElementById("progress-container").style.display = "block"; document.getElementById("progress-bar").style.width = width + "%"; }

// 核心生成
async function handleGenerate() {
    const btn = document.getElementById("generate-btn");
    const selectedItem = document.querySelector(".topic-item.selected");
    const finalTopic = selectedItem ? selectedItem.innerText : document.getElementById("topic").value;

    clearStatus();
    if (!finalTopic || !document.getElementById("course").value) { alert(" 错误：请填写【课程名称】与【章节课题】"); return; }
    if (!selectedClass) { alert(" 错误：请选择或新建【授课班级】！"); return; }
    if (!CONFIG.teacher || !CONFIG.major || !CONFIG.url) { alert(" 错误：系统未就绪！\n请先配置您的信息及API。"); return; }

    const formData = {
        teacher: CONFIG.teacher, major: CONFIG.major, classStr: selectedClass,
        course: document.getElementById("course").value, topic: finalTopic,
        courseType: document.querySelector('input[name="courseType"]:checked').value,
        contentInfo: document.getElementById("content").value || "无"
    };

    try {
        btn.disabled = true; btn.innerText = "生成中..."; updateProgress(10);
        const modelNameDisplay = CONFIG.model || "本地大模型"; showStatus(`正在请求 ${modelNameDisplay}...`, "info");

        let aiResponseJSON = await fetchLessonPlanFromAI(formData);
        if (formData.courseType === "理论课") { aiResponseJSON.practical_content = ""; aiResponseJSON.practical_equipment = ""; }

        updateProgress(70); showStatus(`正在等待 ${modelNameDisplay} 响应...`, "info");
        await writeToWord(formData, aiResponseJSON);

        updateProgress(100); showStatus("教案排版写入成功", "success");
        setTimeout(() => { document.getElementById("progress-container").style.display = "none"; }, 1000);
    } catch (error) {
        document.getElementById("progress-bar").style.background = "var(--error-color)"; showStatus(`错误：${error.message}`, "error");
    } finally {
        btn.disabled = false; setTimeout(() => { btn.innerText = "执行生成并写入文档"; document.getElementById("progress-bar").style.background = ""; }, 2500);
    }
}

async function fetchLessonPlanFromAI(data) {
    const jsonFormatStr = `{"objectives":"目的","practical_content":"内容","practical_equipment":"设备","focus":"重点","difficulties":"难点","aids":"辅助","process_org":"组织","process_new":"新课","process_summary":"小结","process_hw":"作业","postscript":"后记"}`;

    let systemPrompt = `身份：上海中等职业技术学校讲师。专业【${data.major}】、课程【${data.course}】、课题【${data.topic}】。
【极其严格的系统规则（违反将导致程序崩溃）】：
1. 仅输出单一纯JSON对象。所有的值必须是纯文本字符串。绝对禁止包含任何 Markdown 格式符号。
2. 【教学目的】：绝对不能超过 25 个字。
3. 【组织教学】：写授课方式，比如小游戏，举例子。
4. 【讲授新课】：字数控制在 200-350 字左右。必须直接从“1. [小标题名称]”开始，同时要有2，3小标题内容，授课流程要连贯易懂。绝对禁止生成总标题！绝对禁止在结尾输出“字数：xxx”之类的信息。必须包含生活化比喻。
5. 【教学后记】：字数不能超过 50 个字。这是写给教务处看的反思，不要有建议字眼（如：学生掌握良好，下次需加强实物演示）。
6. 【换行规则】：文本内需要换行处，必须输出明文 \\n，绝对禁止物理回车。
7.  教学重点与教学难点要有1，2，3分类，每个分类不超过15个字。教学重点至少有3个，教学难点至少有1个，如果课题难度较高就多写几个。
8.  归纳小结是对整节课讲的内容进行重点总结，不超过50字，但也不要太简单（例如：本节课我们讲解了....）
9.  教学辅助手段可以有PPT课件，3D拆解视频，操作视频。每个词中间用逗号分割
10. 布置作业要与当节课授课内容有关，简单一点的作业。
返回结构体：${jsonFormatStr}`;

    let messagesPayload = [];
    if (uploadedImages.length > 0) {
        let imgInstructions = `\n\n【排版配图特殊指令】：\n教师已在前端上传了 ${uploadedImages.length} 张辅助教学图片：\n`;
        uploadedImages.forEach((img, i) => { imgInstructions += `${i + 1}. 描述："${img.desc || '未命名配图'}" (占位符代码：{{${img.id}}})\n`; });
        imgInstructions += `请在你生成的【讲授新课】段落中，在你认为需要配图的位置，独占一行输出对应的占位符代码！`;
        systemPrompt += imgInstructions;

        let contentArray = [{ type: "text", text: systemPrompt }];
        uploadedImages.forEach(img => {
            let pureBase64 = img.base64.split(',')[1];
            contentArray.push({ type: "image_url", image_url: { url: `data:image/jpeg;base64,${pureBase64}` } });
        });
        messagesPayload = [{ role: "user", content: contentArray }];
    } else {
        messagesPayload = [{ role: "user", content: systemPrompt }];
    }

    const headers = { "Content-Type": "application/json" };
    if (CONFIG.key) headers["Authorization"] = `Bearer ${CONFIG.key}`;

    let response;
    try { response = await fetch(CONFIG.url, { method: "POST", headers: headers, body: JSON.stringify({ model: CONFIG.model, messages: messagesPayload, temperature: 0.7, max_tokens: 4096 }) }); }
    catch (e) { throw new Error("连接失败，请检查 Base URL"); }
    if (!response.ok) throw new Error(`服务端异常 (HTTP ${response.status})`);

    let rawText = (await response.json()).choices[0].message.content;

    let startIndex = rawText.indexOf('{');
    let endIndex = rawText.lastIndexOf('}');
    if (startIndex !== -1 && endIndex !== -1) {
        rawText = rawText.substring(startIndex, endIndex + 1);
    }


    rawText = rawText.replace(/```json/g, '').replace(/```/g, '').trim();

    try {
        return JSON.parse(rawText);
    } catch (e) {

        let safeText = rawText.replace(/\n/g, "\\n").replace(/\\n\{/g, "{\n").replace(/\\n\}/g, "\n}").replace(/",\\n/g, "\",\n").replace(/\\n"/g, "\n\"");
        try {
            return JSON.parse(safeText);
        } catch (e2) {
            throw new Error("模型输出了损坏的 JSON 数据");
        }
    }
}

async function writeToWord(formData, aiData) {
    return Word.run(async (context) => {
        const today = new Date();
        const dateStr = `${today.getFullYear()}年${String(today.getMonth() + 1).padStart(2, '0')}月${String(today.getDate()).padStart(2, '0')}日`;
        const contentMapping = {
            "cc_teacher": formData.teacher || "", "cc_date": dateStr, "cc_major": formData.major || "",
            "cc_class": formData.classStr || "", "cc_course": formData.course || "",
            "cc_topic": formData.topic || "", "cc_course_type": formData.courseType.replace("课", "") || "",
            "cc_objectives": String(aiData.objectives || ""), "cc_prac_content": String(aiData.practical_content || ""),
            "cc_prac_equip": String(aiData.practical_equipment || ""), "cc_focus": String(aiData.focus || ""),
            "cc_difficulties": String(aiData.difficulties || ""), "cc_aids": String(aiData.aids || ""),
            "cc_process_org": String(aiData.process_org || ""), "cc_process_new": String(aiData.process_new || ""),
            "cc_process_summary": String(aiData.process_summary || ""), "cc_process_hw": String(aiData.process_hw || ""),
            "cc_postscript": String(aiData.postscript || "")
        };

        const allControls = context.document.contentControls;
        allControls.load("items/tag, items/title");
        await context.sync();

        let writeCount = 0;
        for (const [key, rawText] of Object.entries(contentMapping)) {
            const targetControls = allControls.items.filter(c => c.tag === key || c.title === key);
            let cleanText = rawText.replace(/\*\*/g, "").replace(/\*/g, "").replace(/#/g, "").replace(/（字数：.*?）/g, "").replace(/\(字数[：:].*?\)/g, "");
            cleanText = cleanText.replace(/\n\s*\n+/g, "\n").trim();

            for (const targetControl of targetControls) {
                if (key === "cc_process_new" && uploadedImages.length > 0) {
                    let parts = cleanText.split(/({{IMG_[a-zA-Z0-9]+}})/).filter(p => p !== "");
                    let isFirst = true;
                    if (parts.length === 0) targetControl.insertText("", "Replace");

                    for (let part of parts) {
                        let match = part.match(/^{{(IMG_[a-zA-Z0-9]+)}}$/);
                        if (match) {
                            let imgObj = uploadedImages.find(i => i.id === match[1]);
                            if (imgObj) {
                                let pureBase64 = imgObj.base64.split(',')[1];
                                let pic = targetControl.insertInlinePictureFromBase64(pureBase64, isFirst ? "Replace" : "End");
                                pic.lockAspectRatio = true; pic.width = 220;
                            }
                        } else {
                            targetControl.insertText(part, isFirst ? "Replace" : "End");
                        }
                        isFirst = false;
                    }
                } else {
                    if (cleanText === "") targetControl.insertText("", "Replace");
                    else targetControl.insertText(cleanText, "Replace");
                }
                writeCount++;
            }
        }
        await context.sync();
        if (writeCount === 0) throw new Error("未检测到对应的占位控件");
    }).catch(e => {
        if (e instanceof OfficeExtension.Error && e.code === "AccessDenied") throw new Error("文档只读，无法写入");
        throw e;
    });
}