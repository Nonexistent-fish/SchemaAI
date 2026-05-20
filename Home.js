Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        injectCustomStyles();
        initStorage();
        initUIEvents();
        initDragToLoad();
    }
});

function injectCustomStyles() {
    if (document.getElementById("schemaai-custom-styles")) return;
    const style = document.createElement("style");
    style.id = "schemaai-custom-styles";
    style.innerHTML = `
        .image-item .img-delete {
            opacity: 0;
            background-color: #ffffff;
            transition: all 0.2s ease;
            border-radius: 50%;
            cursor: pointer;
        }
        .image-item:hover .img-delete {
            opacity: 1;
        }
        .image-item .img-delete:hover {
            background-color: #ff3b30 !important;
        }
        .image-item .img-delete:hover svg line {
            stroke: #ffffff !important;
        }
    `;
    document.head.appendChild(style);
}

let CONFIG = { teacher: "", major: "", model: "", url: "", key: "", classes: [], defaultView: "view-home" };
let selectedClass = "";
let currentTopicCount = 3;
let generatedTopics = [];
const ITEM_HEIGHT = 43;
let uploadedImages = [];
let lastRootView = "view-home";

const AVAILABLE_VIEWS = [
    { id: "view-home", name: "主页" },
    { id: "view-main", name: "教案生成页面" }
];

function showStatus(text, type = "info") { document.getElementById("status-msg").innerHTML = `<span class="status-${type}">${text}</span>`; }
function clearStatus() { document.getElementById("status-msg").innerHTML = ""; }

function switchView(viewId) {
    if (viewId === "view-home" || viewId === "view-main") lastRootView = viewId;
    document.querySelectorAll(".view-section").forEach(el => el.classList.remove("active"));
    document.getElementById(viewId).classList.add("active");
}

function initStorage() {
    const stored = localStorage.getItem("schema_config");
    if (stored) {
        CONFIG = { ...CONFIG, ...JSON.parse(stored) };
        if (CONFIG.classes) CONFIG.classes = CONFIG.classes.filter(c => c && c.trim() !== "");
    } else {
        CONFIG.model = ""; CONFIG.url = ""; CONFIG.classes = []; CONFIG.defaultView = "view-home";
    }

    const viewSelect = document.getElementById("setting-default-view");
    viewSelect.innerHTML = "";
    AVAILABLE_VIEWS.forEach(v => {
        let opt = document.createElement("option");
        opt.value = v.id; opt.innerText = v.name;
        viewSelect.appendChild(opt);
    });

    viewSelect.value = CONFIG.defaultView;
    if (AVAILABLE_VIEWS.find(v => v.id === CONFIG.defaultView)) switchView(CONFIG.defaultView);
    else switchView("view-home");

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

    document.getElementById("ai-topic-btn").onclick = () => {
        const courseInput = document.getElementById("course");
        const topicInput = document.getElementById("topic");

        clearStatus();
        if (!courseInput.value) {
            courseInput.parentElement.classList.add("error-flash");
            setTimeout(() => courseInput.parentElement.classList.remove("error-flash"), 400);
            showStatus("提示：请先填写【课程名称】", "error");
            return;
        }
        if (!topicInput.value) {
            document.getElementById("topic-wrapper").classList.add("error-flash");
            setTimeout(() => document.getElementById("topic-wrapper").classList.remove("error-flash"), 400);
            showStatus("提示：请先填写【章节课题】大类", "error");
            return;
        }
        fetchAndRenderTopics(3, false);
    };

    document.getElementById("topic").addEventListener("input", () => {
        document.querySelectorAll(".topic-item").forEach(el => el.classList.remove("selected"));
    });

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
    document.getElementById("clipboard-img-btn").onclick = async () => {
        try {
            const items = await navigator.clipboard.read();
            let found = false;
            for (const item of items) {
                const types = item.types.filter(t => t.startsWith('image/'));
                for (const type of types) {
                    const blob = await item.getType(type);
                    const reader = new FileReader();
                    reader.onload = (e) => {
                        uploadedImages.push({
                            id: "IMG_" + Math.random().toString(36).substr(2, 6),
                            base64: e.target.result,
                            desc: ""
                        });
                        renderImageGallery();
                    };
                    reader.readAsDataURL(blob);
                    found = true;
                }
            }
            if (!found) showStatus("剪贴板中未检测到图片", "warn");
        } catch (err) {
            showStatus("请授予剪贴板权限，或直接按 Ctrl+V 粘贴", "error");
        }
    };
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

    document.addEventListener('paste', (e) => {
        if (!e.clipboardData) return;
        const items = e.clipboardData.items;
        for (let i = 0; i < items.length; i++) {
            if (items[i].type.indexOf('image') !== -1) {
                const blob = items[i].getAsFile();
                const reader = new FileReader();
                reader.onload = (event) => {
                    uploadedImages.push({ id: "IMG_" + Math.random().toString(36).substr(2, 6), base64: event.target.result, desc: "" });
                    renderImageGallery();
                    showStatus("已成功粘贴图片", "success");
                };
                reader.readAsDataURL(blob);
            }
        }
    });

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

async function fetchAndRenderTopics(targetCount, isDragging) {
    const topicInput = document.getElementById("topic");
    const courseInput = document.getElementById("course").value;
    const panel = document.getElementById("ai-topic-panel");
    const list = document.getElementById("topic-list");
    const aiBtn = document.getElementById("ai-topic-btn");

    if (!CONFIG.url) {
        const setBtn = document.getElementById("nav-settings");
        setBtn.classList.add("error-flash");
        setTimeout(() => setBtn.classList.remove("error-flash"), 400);
        showStatus("提示：未配置 API 接口地址", "error");
        return;
    }

    currentTopicCount = targetCount;
    panel.style.display = "block";
    aiBtn.classList.add("disabled");
    list.style.height = `${targetCount * ITEM_HEIGHT}px`;
    list.innerHTML = `<div class="thinking-center">Thinking...</div>`;

    try {
        const prompt = `课程：${courseInput}。大类：${topicInput.value}。请拆分为正好 ${targetCount} 个课时标题。难度从易到难，让学生容易理解学习。只返回一个纯JSON数组，例如 ["课题1", "课题2", "课题3"]。绝对禁止包含任何解释、前缀（如课题1：）、Markdown 或额外文本。`;
        const headers = { "Content-Type": "application/json" };
        if (CONFIG.key) headers["Authorization"] = `Bearer ${CONFIG.key}`;

        const isDeepSeek = /^deepseek(-v4)?-(flash|pro)$/i.test(CONFIG.model);
        const requestBody = {
            model: CONFIG.model,
            messages: [{ role: "user", content: prompt }],
            temperature: 0.3
        };
        if (isDeepSeek) {
            requestBody.thinking = { type: "disabled" };
        }

        const response = await fetch(CONFIG.url, {
            method: "POST",
            headers: headers,
            body: JSON.stringify(requestBody)
        });
        if (!response.ok) {
            const errText = await response.text().catch(() => "无返回");
            throw new Error(`接口返回错误 (${response.status}): ${errText}`);
        }

        let rawText = (await response.json()).choices[0].message.content;
        rawText = rawText.replace(/<think>[\s\S]*?<\/think>/g, '').trim();
        rawText = rawText.replace(/```json|```/gi, '').trim();

        if (rawText.startsWith('{')) {
            const obj = JSON.parse(rawText);
            const array = obj.topics || obj.items || obj.courses || Object.values(obj).find(Array.isArray);
            if (Array.isArray(array)) {
                rawText = JSON.stringify(array);
            } else {
                throw new Error("返回的JSON对象中未找到数组");
            }
        }

        if (!rawText.startsWith('[')) {
            const arrMatch = rawText.match(/\[([\s\S]*)\]/);
            if (arrMatch) {
                rawText = '[' + arrMatch[1] + ']';
            } else {
                throw new Error("未找到有效的JSON数组");
            }
        }

        generatedTopics = JSON.parse(rawText);
        if (!Array.isArray(generatedTopics)) throw new Error("解析结果不是数组");

        list.innerHTML = "";
        generatedTopics.forEach((text, index) => {
            const item = document.createElement("div");
            item.className = "topic-item placeholder";
            setTimeout(() => {
                item.className = "topic-item slide-in selected";
                if (index > 0) item.classList.remove("selected");
                item.style.animationDelay = `${index * 0.05}s`;
                item.innerText = text.replace(/\*\*/g, "").replace(/\*/g, "");
            }, 50);
            item.onclick = () => {
                document.querySelectorAll(".topic-item").forEach(el => el.classList.remove("selected"));
                item.classList.add("selected");
            };
            list.appendChild(item);
        });
        setTimeout(() => { list.style.height = "auto"; }, 300);
    } catch (error) {
        const errMsg = error.message || error.toString() || "未知错误";
        list.innerHTML = `<div class="thinking-center" style="color:var(--error-color);">失败: ${errMsg}</div>`;
    } finally {
        aiBtn.classList.remove("disabled");
    }
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

async function handleGenerate() {
    const btn = document.getElementById("generate-btn");
    const selectedItem = document.querySelector(".topic-item.selected");
    const finalTopic = selectedItem ? selectedItem.innerText : document.getElementById("topic").value;

    clearStatus();
    if (!document.getElementById("course").value) {
        document.getElementById("course").parentElement.classList.add("error-flash");
        setTimeout(() => document.getElementById("course").parentElement.classList.remove("error-flash"), 400);
        showStatus("阻断：请填写【课程名称】", "error");
        return;
    }
    if (!finalTopic) {
        document.getElementById("topic-wrapper").classList.add("error-flash");
        setTimeout(() => document.getElementById("topic-wrapper").classList.remove("error-flash"), 400);
        showStatus("阻断：请填写或生成【章节课题】", "error");
        return;
    }
    if (!selectedClass) {
        document.getElementById("class-select-wrapper").classList.add("error-flash");
        setTimeout(() => document.getElementById("class-select-wrapper").classList.remove("error-flash"), 400);
        showStatus("阻断：请选择或新建【授课班级】", "error");
        return;
    }
    if (!CONFIG.teacher || !CONFIG.major) {
        const userBtn = document.getElementById("nav-user");
        userBtn.classList.add("error-flash");
        setTimeout(() => userBtn.classList.remove("error-flash"), 400);
        showStatus("阻断：请点击头像设置【教师与专业】", "error");
        return;
    }
    if (!CONFIG.url) {
        const setBtn = document.getElementById("nav-settings");
        setBtn.classList.add("error-flash");
        setTimeout(() => setBtn.classList.remove("error-flash"), 400);
        showStatus("阻断：请点击齿轮配置【API接口】", "error");
        return;
    }

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

        updateProgress(70); showStatus(`等待 ${modelNameDisplay} 响应...`, "info");
        await writeToWord(formData, aiResponseJSON);

        updateProgress(100); showStatus("教案排版写入成功", "success");
        setTimeout(() => { document.getElementById("progress-container").style.display = "none"; }, 1000);
    } catch (error) {
        document.getElementById("progress-bar").style.background = "var(--error-color)"; showStatus(`错误：${error.message}`, "error");
    } finally {
        btn.disabled = false; setTimeout(() => { btn.innerText = "执行生成"; document.getElementById("progress-bar").style.background = ""; }, 2500);
    }
}

async function fetchLessonPlanFromAI(data) {
    const jsonFormatStr = `{"objectives":"目的","practical_content":"内容","practical_equipment":"设备","focus":"重点","difficulties":"难点","aids":"辅助","process_org":"组织","process_new":"新课","process_summary":"小结","process_hw":"作业","postscript":"后记"}`;

    let rolePrompt = `身份：上海中等职业技术学校讲师。专业${data.major}、课程${data.course}、课题${data.topic}。`;
    if (data.courseType === "实训课") {
        rolePrompt = `身份：上海中等职业技术学校【实训指导高级教师】。正在实训场地上实训操作课。专业${data.major}、实操课题${data.topic}。`;
    }

    let systemPrompt = `${rolePrompt}
请务必遵守以下规则：
1. 仅输出单一纯JSON对象。所有的值必须是纯文本。绝对禁止包含 Markdown，不允许输出除JSON对象外其他多余语句。
2. 教学目的：不能超过 25 字。
3. 组织教学：写授课方式，如小游戏，举例子。
4. 讲授新课：300-450字。必须从“1. 小标题名称”开始，包含2-5小标题，包含生活化比喻。禁止生成总标题和结尾字数统计！
5. 教学后记：不超过50字的反思，禁止建议字眼。
6. 换行规则：需要换行处输出明文 \\n，绝对禁止物理回车。
7.  教学重点与难点要有1，2，3序号排序，每个序号独占一行，禁止用逗号分隔，每项不超过15字。重点至少3个，难点至少1个。
8.  归纳小结是对内容的重点总结，不超过50字。
9.  教学辅助手段逗号分割。
10. 布置简单相关作业,不要生成序号`;

    if (data.courseType === "实训课") {
        systemPrompt += `\n11. 实训指令：讲授新课和组织教学必须围绕“实物拆装、设备操作”展开。绝对禁止纯理论！必须设计学生分组实操环节！`;
    }

    systemPrompt += `\n返回结构：${jsonFormatStr}`;

    const headers = { "Content-Type": "application/json" };
    if (CONFIG.key) headers["Authorization"] = `Bearer ${CONFIG.key}`;

    const isDeepSeekV4 = /^deepseek(-v4)?-(flash|pro)$/i.test(CONFIG.model);

    let messagesPayload = [];
    if (uploadedImages.length > 0) {
        // 包含图片base64 + 让模型自行描述图片
        let imgInstructions = `\n\n排版配图指令：\n前端已上传 ${uploadedImages.length} 张图片。请仔细观察每张图片，并为需要配图的讲授新课内容，在合适位置独占一行输出对应的占位符代码：\n`;
        uploadedImages.forEach((img, i) => {
            imgInstructions += `${i + 1}. 第${i + 1}张图片 (占位符：{{${img.id}}})\n`;
        });
        imgInstructions += `请在讲授新课需要配图处，独占一行输出对应的占位符代码！`;
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
    async function makeRequest(payload, description = "") {
        const requestBody = {
            model: CONFIG.model,
            messages: payload,
            max_tokens: 4096
        };
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
            throw new Error(`服务端异常 (HTTP ${response.status})${description ? ': ' + description : ''}`);
        }
        let rawText = (await response.json()).choices[0].message.content;
        return rawText;
    }

    function parseResponse(rawText) {
        rawText = rawText.replace(/<think>[\s\S]*?<\/think>/g, '').trim();
        rawText = rawText.replace(/```json/gi, '').replace(/```/g, '').trim();

        if (rawText.startsWith('[')) rawText = rawText.replace(/^\[/, '').replace(/\]$/, '').trim();

        let startIndex = rawText.indexOf('{');
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

            let endIndex = rawText.lastIndexOf('}');
            if (endIndex !== -1) {
                rawText = rawText.substring(0, endIndex + 1);
            }
        }

        try {
            return JSON.parse(rawText);
        } catch (e) {
            let safeText = rawText.replace(/\n/g, "\\n").replace(/\r/g, "");
            safeText = safeText.replace(/,\s*\}/g, '}');
            try {
                return JSON.parse(safeText);
            } catch (e2) {
                throw new Error("模型输出了严重损坏的数据，请更换模型重试");
            }
        }
    }

    try {

        const rawText = await makeRequest(messagesPayload);
        return parseResponse(rawText);
    } catch (error) {

        if (uploadedImages.length === 0) {
            throw error; 
        }

        const missingDescs = uploadedImages.some(img => !img.desc || img.desc.trim() === '');
        if (missingDescs) {
            throw new Error("当前模型可能不支持多模态识别，请点击图片填写“图片简述”，填写后工具将自动在文本中插入图片。");
        }

        let fallbackPrompt = systemPrompt; 
        let imgInstructions = `\n\n排版配图指令（无法看到图片，根据描述决策）：\n前端已上传 ${uploadedImages.length} 张图片，它们的描述如下，请根据描述在讲授新课中合适的位置独占一行插入对应的占位符代码。\n`;
        uploadedImages.forEach((img, i) => {
            imgInstructions += `${i + 1}. 描述："${img.desc || '未命名配图'}" (占位符：{{${img.id}}})\n`;
        });
        imgInstructions += `请在讲授新课需要配图处，独占一行输出对应的占位符代码！`;
        fallbackPrompt += imgInstructions;
        fallbackPrompt += `\n返回结构：${jsonFormatStr}`;

        const fallbackPayload = [{ role: "user", content: fallbackPrompt }];
        try {
            const rawText = await makeRequest(fallbackPayload, "降级为纯文本模式后");
            return parseResponse(rawText);
        } catch (retryError) {
            throw new Error(`图片上传失败: ${retryError.message}`);
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

        function isTitleLine(line) {
            const trimmed = line.trim();
            // 匹配：数字开头 + 可选空格 + 标题符号（点、顿号、括号等）
            return /^\d+\s*[\.．、)）]/.test(trimmed);
        }

        const allControls = context.document.contentControls;
        allControls.load("items/tag, items/title");
        await context.sync();

        let writeCount = 0;
        let imgCounter = 1;
        const targetKeys = ["cc_process_org", "cc_process_new", "cc_process_summary", "cc_process_hw", "cc_course_type"];
        const headerKeys = ["cc_teacher", "cc_date", "cc_major", "cc_class", "cc_course", "cc_topic"];

        for (const [key, rawText] of Object.entries(contentMapping)) {
            if (!rawText || rawText.trim() === "") continue;

            const targetControls = allControls.items.filter(c => c.tag === key || c.title === key);
            let cleanText = rawText.replace(/\*\*/g, "").replace(/\*/g, "").replace(/#/g, "").replace(/（字数：.*?）/g, "").replace(/\(字数[：:].*?\)/g, "");
            cleanText = cleanText.replace(/\n\s*\n+/g, "\n").trim();

            for (const targetControl of targetControls) {
                let isFirstInsert = true;

                // 图文混排（图片占位符处理）
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
                                targetControl.insertText("\n　　　　　　", "End");

                                let pic = targetControl.insertInlinePictureFromBase64(pureBase64, "End");
                                pic.lockAspectRatio = true;
                                pic.width = 220;

                                let descriptionStr = imgObj.desc || "自动变速箱配件";
                                let captionText = `\n　　　　　　图${imgCounter}  ${descriptionStr}\n`;
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
                            let lines = part.split('\n');
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
                    let lines = cleanText.split('\n');
                    for (let line of lines) {
                        if (line.trim() !== "") {
                            let insertStr = isFirstInsert ? line.trim() : "\n" + line.trim();
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
        if (writeCount === 0) throw new Error("未检测到对应的占位控件");
    }).catch(e => {
        if (e instanceof OfficeExtension.Error && e.code === "AccessDenied") throw new Error("文档只读，无法写入");
        throw e;
    });
}