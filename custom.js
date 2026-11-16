// ========== 自定义模板双端分发与使用界面渲染 ==========

function refreshCustomView() {
    if (window.currentHost === Office.HostType.Word) {
        if (typeof refreshWordCustomView === 'function') refreshWordCustomView();
    } else if (window.currentHost === Office.HostType.Excel) {
        if (typeof refreshExcelCustomView === 'function') refreshExcelCustomView();
    }
}

function renderTemplateUseView() {
    const template = CONFIG.customTemplateData;
    if (!template) return;

    const titleEl = document.getElementById("template-use-title");
    const descEl = document.getElementById("template-use-desc");
    const container = document.getElementById("template-use-controls-list");

    if (titleEl) titleEl.innerText = template.title || "模板";
    if (descEl) descEl.innerText = template.desc || "";
    if (!container) return;

    container.innerHTML = "";
    const config = template.controlsConfig || {};
    const titles = template.titles || {};

    Object.keys(config).forEach(tag => {
        const cfg = config[tag];
        if (cfg.inputType === 'none') return;

        const itemDiv = document.createElement('div');
        itemDiv.className = 'template-use-item';
        itemDiv.dataset.tag = tag;

        const headerDiv = document.createElement('div');
        headerDiv.className = 'control-header';
        headerDiv.style.cursor = 'default';
        const displayTitle = titles[tag] || tag;
        headerDiv.innerHTML = `
            <div class="control-header-left">
                <span style="font-weight:600;font-size:14px;">${displayTitle}</span>
                <span class="control-tag">${cfg.targetRange || tag}</span>
            </div>
        `;
        itemDiv.appendChild(headerDiv);

        if (cfg.inputType === 'manual') {
            if (cfg.manualSubType === 'text') {
                const input = document.createElement('input');
                input.type = 'text';
                input.className = 'template-use-input';
                input.placeholder = cfg.displayText || '请输入文本';
                itemDiv.appendChild(input);
            } else if (cfg.manualSubType === 'boolean') {
                const label = document.createElement('label');
                label.className = 'switch';
                label.style.marginTop = '10px';
                const checkbox = document.createElement('input');
                checkbox.type = 'checkbox';
                checkbox.checked = (cfg.booleanDefault === 'true');
                label.appendChild(checkbox);
                const slider = document.createElement('span');
                slider.className = 'slider round';
                label.appendChild(slider);
                itemDiv.appendChild(label);
            } else if (cfg.manualSubType === 'choice') {
                const options = cfg.choiceOptions || [];
                if (cfg.choiceType === 'dropdown') {
                    const selectEl = document.createElement('select');
                    selectEl.className = 'template-use-input';
                    options.forEach((opt, idx) => {
                        const option = document.createElement('option');
                        option.value = idx;
                        option.textContent = opt;
                        if (idx === cfg.choiceDefaultIndex) option.selected = true;
                        selectEl.appendChild(option);
                    });
                    itemDiv.appendChild(selectEl);
                } else {
                    options.forEach((opt, idx) => {
                        const row = document.createElement('div');
                        row.className = 'choice-option-row';
                        row.style.marginTop = '8px';
                        const input = document.createElement('input');
                        if (cfg.choiceType === 'multiple') {
                            input.type = 'checkbox';
                            input.checked = (cfg.choiceDefaultIndexes || []).includes(idx);
                        } else {
                            input.type = 'radio';
                            input.name = `use-choice-${tag}`;
                            input.checked = (cfg.choiceDefaultIndex === idx);
                        }
                        const text = document.createElement('span');
                        text.textContent = opt;
                        text.style.fontSize = '13px';
                        row.appendChild(input);
                        row.appendChild(text);
                        itemDiv.appendChild(row);
                    });
                }
            } else if (cfg.manualSubType === 'image' && window.currentHost === Office.HostType.Word) {
                if (!cfg.images) cfg.images = [];
                const galleryDiv = document.createElement('div');
                galleryDiv.className = 'image-gallery';
                const renderGallery = () => {
                    galleryDiv.innerHTML = '';
                    cfg.images.forEach((img, idx) => {
                        const imgDiv = document.createElement('div');
                        imgDiv.className = 'image-item';
                        imgDiv.innerHTML = `
                            <img src="${img.base64}" class="img-preview" title="${img.desc || '点击添加图片描述'}">
                            ${img.desc ? '<div class="img-has-desc" title="已添加简述"></div>' : ''}
                            <div class="img-delete" data-index="${idx}"><svg viewBox="0 0 24 24"><line x1="18" y1="6" x2="6" y2="18"></line><line x1="6" y1="6" x2="18" y2="18"></line></svg></div>
                        `;
                        imgDiv.querySelector('.img-delete').onclick = (e) => {
                            e.stopPropagation();
                            cfg.images.splice(idx, 1);
                            renderGallery();
                        };
                        galleryDiv.appendChild(imgDiv);
                    });
                };
                renderGallery();

                const addBtn = document.createElement('button');
                addBtn.className = 'icon-btn small';
                addBtn.style.marginTop = '8px';
                addBtn.innerHTML = `<svg viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"></line><line x1="5" y1="12" x2="19" y2="12"></line></svg>`;
                addBtn.title = '导入本地图片';
                addBtn.onclick = () => {
                    const input = document.createElement('input');
                    input.type = 'file';
                    input.accept = 'image/*';
                    input.onchange = (e) => {
                        Array.from(e.target.files).forEach(file => {
                            const reader = new FileReader();
                            reader.onload = (ev) => {
                                cfg.images.push({ base64: ev.target.result, caption: '', desc: '' });
                                renderGallery();
                            };
                            reader.readAsDataURL(file);
                        });
                    };
                    input.click();
                };
                itemDiv.appendChild(galleryDiv);
                itemDiv.appendChild(addBtn);
            }
        }
        container.appendChild(itemDiv);
    });

    const bar = document.createElement("div");
    bar.className = "generate-action-bar";
    bar.innerHTML = `
        <button id="template-generate-btn" class="primary-btn" style="margin-top:0;">执行生成</button>
        <button id="template-save-doc-btn" class="save-icon-btn" title="一键保存文档">
            <svg viewBox="0 0 24 24" width="20" height="20" stroke="currentColor" stroke-width="2" fill="none"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
        </button>
    `;
    container.appendChild(bar);

    const progressContainer = document.createElement('div');
    progressContainer.className = 'progress-container';
    progressContainer.style.marginTop = '12px';
    progressContainer.innerHTML = '<div class="progress-bar" style="width:0%"></div>';
    container.appendChild(progressContainer);

    const statusMsg = document.createElement('div');
    statusMsg.className = 'status-msg';
    container.appendChild(statusMsg);

    const genBtn = document.getElementById("template-generate-btn");
    if (genBtn) {
        genBtn.addEventListener("click", async () => {
            const btn = document.getElementById("template-generate-btn");
            const saveBtn = document.getElementById("template-save-doc-btn");
            const progressBar = progressContainer.querySelector('.progress-bar');
            const updateProgress = (w) => { progressContainer.style.display = 'block'; if (progressBar) progressBar.style.width = w + '%'; };
            const localShowStatus = (t, type = 'info') => { statusMsg.innerHTML = `<span class="status-${type}">${t}</span>`; };

            if (saveBtn) saveBtn.classList.remove("show");
            const userInputs = collectCustomUserInputs();

            // 判断是否有任何控件启用了 AI
            const configObj = template.controlsConfig || {};
            const needsAI = Object.values(configObj).some(c => c.aiEnabled || c.trueAction === 'ai' || c.falseAction === 'ai');

            if (needsAI && !CONFIG.url) {
                localShowStatus("请先配置 API 接口地址", "error");
                return;
            }

            try {
                if (btn) { btn.disabled = true; btn.innerText = "生成中..."; }
                updateProgress(20);

                let aiData = Object.assign({}, userInputs);
                if (needsAI) {
                    localShowStatus(`正在请求 ${CONFIG.model || 'AI'}...`, "info");
                    const fetched = await fetchCustomAIResponse(template, userInputs);
                    aiData = Object.assign({}, userInputs, fetched);
                }

                updateProgress(70);
                localShowStatus(`正在写入文档...`, "info");

                if (window.currentHost === Office.HostType.Word) {
                    await writeWordCustomData(template, aiData);
                } else if (window.currentHost === Office.HostType.Excel) {
                    await writeExcelCustomData(template, aiData);
                }

                updateProgress(100);
                localShowStatus("写入成功", "success");
                setTimeout(() => {
                    progressContainer.style.display = "none";
                    if (saveBtn) saveBtn.classList.add("show");
                }, 1200);
            } catch (error) {
                if (progressBar) progressBar.style.background = "var(--error-color)";
                localShowStatus(`错误：${error.message}`, "error");
            } finally {
                if (btn) {
                    btn.disabled = false;
                    setTimeout(() => { btn.innerText = "执行生成"; if (progressBar) progressBar.style.background = ""; }, 2500);
                }
            }
        });
    }

    const saveDocBtn = document.getElementById("template-save-doc-btn");
    if (saveDocBtn) {
        saveDocBtn.addEventListener("click", () => {
            if (window.currentHost === Office.HostType.Excel && typeof handleExcelSaveDocument === 'function') {
                handleExcelSaveDocument();
            } else if (window.currentHost === Office.HostType.Word && typeof handleSaveDocument === 'function') {
                handleSaveDocument();
            }
        });
    }
}

function collectCustomUserInputs() {
    const inputs = {};
    document.querySelectorAll('.template-use-item').forEach(item => {
        const tag = item.dataset.tag || (item.querySelector('.control-tag') ? item.querySelector('.control-tag').textContent.trim() : '');
        if (!tag) return;

        const textInput = item.querySelector('input[type="text"].template-use-input');
        if (textInput) { inputs[tag] = textInput.value.trim(); return; }

        const switchBox = item.querySelector('.switch input[type="checkbox"]');
        if (switchBox) { inputs[tag] = switchBox.checked ? 'true' : 'false'; return; }

        const radios = item.querySelectorAll('input[type="radio"]');
        if (radios.length > 0) {
            const checked = item.querySelector('input[type="radio"]:checked');
            if (checked) {
                const span = checked.closest('.choice-option-row')?.querySelector('span');
                inputs[tag] = span ? span.textContent.trim() : '';
            }
            return;
        }

        const checkboxes = item.querySelectorAll('.choice-option-row input[type="checkbox"]');
        if (checkboxes.length > 0) {
            const vals = [];
            checkboxes.forEach(cb => {
                if (cb.checked) {
                    const span = cb.closest('.choice-option-row')?.querySelector('span');
                    if (span) vals.push(span.textContent.trim());
                }
            });
            inputs[tag] = vals.join(', ');
            return;
        }

        const selectEl = item.querySelector('select.template-use-input');
        if (selectEl && selectEl.selectedIndex >= 0) {
            inputs[tag] = selectEl.options[selectEl.selectedIndex].text;
        }
    });
    return inputs;
}

async function fetchCustomAIResponse(template, userInputs = {}) {
    const config = template.controlsConfig || {};
    const titles = template.titles || {};
    const tags = Object.keys(config);

    let systemPrompt = `你是专业教学与办公内容生成助手。请根据以下字段与技能要求生成内容，并严格返回一个纯 JSON 对象。\n模板：${template.title} (${template.desc || ''})\n\n字段列表：\n`;

    tags.forEach(tag => {
        const cfg = config[tag];
        const title = titles[tag] || tag;

        // 匹配内置 Skill 预设
        let skillPrompt = "";
        if (cfg.aiSkillId && typeof SKILL_PRESETS !== 'undefined') {
            const matchedSkill = SKILL_PRESETS.find(s => s.id === cfg.aiSkillId);
            if (matchedSkill && matchedSkill.prompt) skillPrompt = matchedSkill.prompt;
        }

        systemPrompt += `- 键名 "${tag}"（标题："${title}"）`;
        if (skillPrompt) systemPrompt += `，技能要求：${skillPrompt}`;
        if (cfg.aiPrompt) systemPrompt += `，补充要求：${cfg.aiPrompt}`;
        if (cfg.manualSubType === 'boolean' && userInputs[tag] !== undefined) {
            const bVal = userInputs[tag];
            if (bVal === 'true' && cfg.trueAction === 'ai' && cfg.truePrompt) systemPrompt += `，特定要求：${cfg.truePrompt}`;
            if (bVal === 'false' && cfg.falseAction === 'ai' && cfg.falsePrompt) systemPrompt += `，特定要求：${cfg.falsePrompt}`;
        }
        if (userInputs[tag]) systemPrompt += `，当前用户输入：${userInputs[tag]}`;
        systemPrompt += `\n`;
    });

    systemPrompt += `\n返回格式：只返回纯 JSON，其键名必须与上述键名严格一致，值禁止包含 Markdown 或多余说明。`;

    const headers = { "Content-Type": "application/json" };
    if (CONFIG.key) headers["Authorization"] = `Bearer ${CONFIG.key}`;

    const response = await fetch(CONFIG.url, {
        method: "POST",
        headers,
        body: JSON.stringify({
            model: CONFIG.model,
            messages: [{ role: "user", content: systemPrompt }],
            temperature: 0.6,
            max_tokens: 8192
        })
    });

    const resText = await response.text();
    if (!response.ok) throw new Error(`AI 请求失败 (${response.status}): ${resText.slice(0, 60)}`);

    let jsonRes;
    try {
        jsonRes = JSON.parse(resText);
    } catch (e) {
        throw new Error("接口返回了非JSON网页内容，请检查API_url是否正确");
    }

    const choice = jsonRes.choices && jsonRes.choices[0];
    let raw = (choice && choice.message) ? (choice.message.content || choice.message.reasoning_content || "") : "";

    raw = raw.replace(/<think>[\s\S]*?<\/think>/g, "").trim().replace(/\x60{3}(?:json)?|\x60{3}/gi, "").trim();
    const start = raw.indexOf("{");
    const end = raw.lastIndexOf("}");
    if (start !== -1 && end !== -1 && end > start) raw = raw.substring(start, end + 1);

    return JSON.parse(raw);
}

function renderSavedTemplateCards() {
    var grid = document.querySelector("#view-home .template-grid");
    if (!grid) return;

    var existingCards = grid.querySelectorAll(".template-card.saved-template");
    existingCards.forEach(function (c) { c.remove(); });

    var hasCustom = (CONFIG.customTemplateData && CONFIG.customTemplateData.controlsConfig && Object.keys(CONFIG.customTemplateData.controlsConfig).length > 0);

    if (hasCustom) {
        var card = document.createElement("div");
        card.className = "template-card saved-template";
        card.style.order = "-1";

        var title = CONFIG.customTemplateData.title || "自定义模板";
        var desc = CONFIG.customTemplateData.desc || "点击使用该模板";

        card.innerHTML =
            "<div style='position:relative;'>" +
            "<div class='card-icon' style='color: var(--success-color);'>" +
            "<svg viewBox='0 0 24 24'><rect x='3' y='3' width='18' height='18' rx='2' ry='2'></rect><line x1='3' y1='9' x2='21' y2='9'></line><line x1='9' y1='21' x2='9' y2='9'></line></svg>" +
            "</div>" +
            "<button class='delete-template-btn' title='删除模板'>" +
            "<svg viewBox='0 0 24 24' width='16' height='16' stroke='currentColor' stroke-width='2' fill='none'><line x1='18' y1='6' x2='6' y2='18'/><line x1='6' y1='6' x2='18' y2='18'/></svg>" +
            "</button>" +
            "</div>" +
            "<h3 style='margin: 0 0 6px 0; font-size: 13px; font-weight: 600; color: var(--text-main);'>" + title + "</h3>" +
            "<p style='margin: 0; font-size: 12px; color: var(--text-sub); line-height: 1.4;'>" + desc + "</p>";

        var delBtn = card.querySelector('.delete-template-btn');
        if (delBtn) {
            delBtn.addEventListener('click', function (e) {
                e.stopPropagation();
                CONFIG.customTemplateData = null;
                saveStorage();
                renderSavedTemplateCards();
                if (typeof initStorageDeferred === 'function') initStorageDeferred();
            });
        }

        card.addEventListener("click", function () {
            if (typeof renderTemplateUseView === 'function') renderTemplateUseView();
            switchView("view-template-use");
        });

        grid.appendChild(card);
    }
}