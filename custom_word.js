// ========== Word 自定义模板专属模块 ==========
let wordControlsConfig = {};
let isWordRefreshing = false;

Office.onReady(() => {
    if (window.currentHost !== Office.HostType.Word) return;

    // 保存模板按钮
    const saveBtn = document.getElementById("save-template-btn");
    if (saveBtn) {
        saveBtn.addEventListener("click", () => {
            if (window.currentHost !== Office.HostType.Word) return;
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
            Object.keys(wordControlsConfig).forEach(tag => {
                cleanConfig[tag] = sanitizeWordConfig(wordControlsConfig[tag]);
                const itemDiv = document.querySelector(`.control-item[data-tag="${tag}"]`);
                if (itemDiv) {
                    const ti = itemDiv.querySelector('.control-title-input');
                    titles[tag] = ti ? ti.value.trim() : '';
                }
            });
            CONFIG.customTemplateData = { title, desc, controlsConfig: cleanConfig, titles, host: "Word" };
            saveStorage();

            showStatus("Word 模板已保存", "success");
            switchView("view-home");
        });
    }

    // 编辑模板按钮
    const editBtn = document.getElementById("edit-template-btn");
    if (editBtn) {
        editBtn.addEventListener("click", () => {
            if (window.currentHost !== Office.HostType.Word) return;
            switchView("view-custom");
            const titleInput = document.getElementById("custom-template-title");
            const descInput = document.getElementById("custom-template-desc");
            if (titleInput) titleInput.value = CONFIG.customTemplateTitle || "";
            if (descInput) descInput.value = CONFIG.customTemplateDesc || "";

            if (CONFIG.customTemplateData && CONFIG.customTemplateData.controlsConfig) {
                const saved = CONFIG.customTemplateData.controlsConfig;
                wordControlsConfig = {};
                Object.keys(saved).forEach(tag => { wordControlsConfig[tag] = sanitizeWordConfig(saved[tag]); });
            } else {
                wordControlsConfig = {};
            }
            refreshWordCustomView();
        });
    }

    // 刷新控件检测
    const refreshBtn = document.getElementById("refresh-custom-btn");
    if (refreshBtn) {
        refreshBtn.addEventListener("click", () => {
            if (window.currentHost === Office.HostType.Word) refreshWordCustomView();
        });
    }
});

function sanitizeWordConfig(cfg) {
    const def = {
        inputType: 'none', manualSubType: 'boolean', aiEnabled: false, aiSkillId: 'custom', aiPrompt: '',
        thinkingMode: false, displayText: '', booleanText: '',
        falseAction: 'none', falseActionFixedText: '', trueAction: 'none', trueActionFixedText: '',
        booleanDefault: 'none', choiceType: 'single',
        choiceOptions: ['选项1', '选项2'], choiceDefaultIndex: 0, choiceDefaultIndexes: [],
        images: [], autoInputText: '', generateCaption: false, truePrompt: '', falsePrompt: ''
    };
    const clean = {};
    Object.keys(def).forEach(key => {
        let value = cfg[key];
        if (value === undefined || value === null) { clean[key] = def[key]; return; }
        if (Array.isArray(def[key])) {
            if (!Array.isArray(value)) clean[key] = def[key];
            else {
                clean[key] = value.map(item => typeof item === 'string' ? item.trim() : item);
                if (key === 'choiceOptions' && clean[key].length < 2) clean[key] = def[key];
            }
        } else if (typeof def[key] === 'boolean') {
            clean[key] = !!value;
        } else if (typeof def[key] === 'number') {
            clean[key] = isNaN(Number(value)) ? def[key] : Number(value);
        } else {
            clean[key] = typeof value === 'string' ? value.trim() : value;
        }
    });
    return clean;
}

async function refreshWordCustomView() {
    if (isWordRefreshing || window.currentHost !== Office.HostType.Word) return;
    isWordRefreshing = true;

    const listContainer = document.getElementById("custom-controls-list");
    const emptyState = document.getElementById("custom-empty-state");
    if (!listContainer || !emptyState) { isWordRefreshing = false; return; }

    listContainer.innerHTML = '';
    listContainer.style.display = 'none';
    emptyState.style.display = 'none';

    try {
        const groups = {};
        await Word.run(async (context) => {
            const contentControls = context.document.contentControls;
            contentControls.load('items/tag, items/title');
            await context.sync();

            const controlsWithTag = contentControls.items.filter(c => c.tag && c.tag.trim() !== '');
            if (controlsWithTag.length === 0) {
                emptyState.style.display = 'block';
                return;
            }

            controlsWithTag.forEach(c => {
                const tag = c.tag.trim();
                if (!groups[tag]) groups[tag] = { tag, title: c.title || '', controls: [] };
                groups[tag].controls.push(c);
            });
        });

        if (Object.keys(groups).length === 0) {
            emptyState.style.display = 'block';
            isWordRefreshing = false;
            return;
        }

        listContainer.style.display = 'block';

        Object.keys(groups).forEach(tag => {
            const group = groups[tag];
            wordControlsConfig[tag] = sanitizeWordConfig(wordControlsConfig[tag] || {});
            const cfg = wordControlsConfig[tag];

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
            tagSpan.textContent = tag;
            headerLeft.appendChild(titleInput);
            headerLeft.appendChild(tagSpan);

            const arrowSvg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
            arrowSvg.setAttribute('viewBox', '0 0 24 24');
            arrowSvg.setAttribute('stroke', 'currentColor');
            arrowSvg.setAttribute('stroke-width', '2');
            arrowSvg.setAttribute('fill', 'none');
            arrowSvg.setAttribute('stroke-linecap', 'round');
            arrowSvg.classList.add('control-arrow');
            const polyline = document.createElementNS('http://www.w3.org/2000/svg', 'polyline');
            polyline.setAttribute('points', '6 9 12 15 18 9');
            arrowSvg.appendChild(polyline);
            headerDiv.appendChild(headerLeft);
            headerDiv.appendChild(arrowSvg);
            headerDiv.addEventListener('click', (e) => {
                if (e.target.tagName === 'INPUT') return;
                itemDiv.classList.toggle('expanded');
            });

            const bodyDiv = document.createElement('div');
            bodyDiv.className = 'control-body';

            const rowDiv = document.createElement('div');
            rowDiv.className = 'control-section-row';
            const inputLabel = document.createElement('label');
            inputLabel.textContent = '输入框内容';
            const inputTypeSelect = document.createElement('select');
            inputTypeSelect.className = 'custom-select-mini input-type-select';
            inputTypeSelect.innerHTML = '<option value="none">无</option><option value="manual">手动输入</option>';
            inputTypeSelect.value = cfg.inputType;
            const typeLabel = document.createElement('label');
            typeLabel.textContent = '类型';
            const manualTypeSelect = document.createElement('select');
            manualTypeSelect.className = 'custom-select-mini manual-sub-type-select';
            manualTypeSelect.innerHTML = `
                <option value="boolean">布尔值</option>
                <option value="choice">选择</option>
                <option value="text">文本</option>
                <option value="image">图片</option>
            `;
            manualTypeSelect.value = cfg.manualSubType;
            const choiceTypeLabel = document.createElement('label');
            choiceTypeLabel.textContent = '选择类型';
            const choiceTypeSelect = document.createElement('select');
            choiceTypeSelect.className = 'custom-select-mini choice-type-select';
            choiceTypeSelect.innerHTML = `
                <option value="single">单选</option>
                <option value="multiple">多选</option>
                <option value="dropdown">下拉菜单</option>
            `;
            choiceTypeSelect.value = cfg.choiceType;
            const showManual = cfg.inputType === 'manual';
            typeLabel.style.display = showManual ? '' : 'none';
            manualTypeSelect.style.display = showManual ? '' : 'none';
            const showChoiceType = showManual && cfg.manualSubType === 'choice';
            choiceTypeLabel.style.display = showChoiceType ? '' : 'none';
            choiceTypeSelect.style.display = showChoiceType ? '' : 'none';
            rowDiv.appendChild(inputLabel);
            rowDiv.appendChild(inputTypeSelect);
            rowDiv.appendChild(typeLabel);
            rowDiv.appendChild(manualTypeSelect);
            rowDiv.appendChild(choiceTypeLabel);
            rowDiv.appendChild(choiceTypeSelect);

            const inputSectionTitle = document.createElement('div');
            inputSectionTitle.className = 'control-section-title';
            inputSectionTitle.textContent = '输入框设置';
            bodyDiv.appendChild(inputSectionTitle);

            const inputSubContainer = document.createElement('div');
            inputSubContainer.className = 'control-sub-item';
            inputSubContainer.appendChild(rowDiv);

            const manualOptionsContainer = document.createElement('div');
            manualOptionsContainer.className = 'manual-options-container';
            inputSubContainer.appendChild(manualOptionsContainer);

            const autoInputGroup = document.createElement('div');
            autoInputGroup.className = 'input-group';
            autoInputGroup.style.display = (cfg.inputType === 'none' && !cfg.aiEnabled) ? '' : 'none';
            const autoLabel = document.createElement('label');
            autoLabel.textContent = '自动输入';
            const autoInput = document.createElement('input');
            autoInput.type = 'text';
            autoInput.className = 'auto-input-text';
            autoInput.value = cfg.autoInputText;
            autoInput.addEventListener('input', () => { cfg.autoInputText = autoInput.value; });
            autoInputGroup.appendChild(autoLabel);
            autoInputGroup.appendChild(autoInput);
            inputSubContainer.appendChild(autoInputGroup);
            bodyDiv.appendChild(inputSubContainer);

            // AI 设置
            const aiSectionTitle = document.createElement('div');
            aiSectionTitle.className = 'control-section-title';
            aiSectionTitle.style.display = 'flex';
            aiSectionTitle.style.alignItems = 'center';
            aiSectionTitle.style.justifyContent = 'space-between';
            const aiTitleText = document.createElement('span');
            aiTitleText.textContent = 'AI 设置';
            const aiSwitch = document.createElement('label');
            aiSwitch.className = 'switch';
            const aiCheckbox = document.createElement('input');
            aiCheckbox.type = 'checkbox';
            aiCheckbox.className = 'ai-checkbox';
            aiCheckbox.checked = cfg.aiEnabled;
            const aiSlider = document.createElement('span');
            aiSlider.className = 'slider round';
            aiSwitch.appendChild(aiCheckbox);
            aiSwitch.appendChild(aiSlider);
            aiSectionTitle.appendChild(aiTitleText);
            aiSectionTitle.appendChild(aiSwitch);
            bodyDiv.appendChild(aiSectionTitle);

            const aiSubContainer = document.createElement('div');
            aiSubContainer.className = 'control-sub-item';

            const aiSettingsDiv = document.createElement('div');
            aiSettingsDiv.className = 'ai-settings';
            aiSettingsDiv.style.display = cfg.aiEnabled ? 'block' : 'none';

            const generalPromptGroup = document.createElement('div');
            generalPromptGroup.className = 'input-group';
            generalPromptGroup.style.marginTop = '8px';
            const generalPromptLabel = document.createElement('label');
            generalPromptLabel.textContent = 'AI提示词';
            const generalPromptInput = document.createElement('input');
            generalPromptInput.type = 'text';
            generalPromptInput.className = 'ai-prompt-input';
            generalPromptInput.placeholder = '输入提示词...';
            generalPromptInput.value = cfg.aiPrompt || '';
            generalPromptInput.addEventListener('input', () => { cfg.aiPrompt = generalPromptInput.value; });
            generalPromptGroup.appendChild(generalPromptLabel);
            generalPromptGroup.appendChild(generalPromptInput);
            aiSettingsDiv.appendChild(generalPromptGroup);

            const thinkRow = document.createElement('div');
            thinkRow.style.display = 'flex';
            thinkRow.style.alignItems = 'center';
            thinkRow.style.gap = '8px';
            thinkRow.style.marginTop = '12px';
            const thinkLabel = document.createElement('label');
            thinkLabel.textContent = '思考模式';
            const thinkSwitch = document.createElement('label');
            thinkSwitch.className = 'switch';
            const thinkCheckbox = document.createElement('input');
            thinkCheckbox.type = 'checkbox';
            thinkCheckbox.className = 'thinking-checkbox';
            thinkCheckbox.checked = cfg.thinkingMode;
            const thinkSlider = document.createElement('span');
            thinkSlider.className = 'slider round';
            thinkSwitch.appendChild(thinkCheckbox);
            thinkSwitch.appendChild(thinkSlider);
            thinkRow.appendChild(thinkLabel);
            thinkRow.appendChild(thinkSwitch);
            aiSettingsDiv.appendChild(thinkRow);

            const booleanPromptsContainer = document.createElement('div');
            booleanPromptsContainer.style.display = 'none';
            booleanPromptsContainer.innerHTML = `
                <div class="input-group" style="margin-top:8px;" data-prompt-type="true">
                    <label>true 提示词</label>
                    <input type="text" class="true-prompt-input" placeholder="true 提示词...">
                </div>
                <div class="input-group" data-prompt-type="false">
                    <label>false 提示词</label>
                    <input type="text" class="false-prompt-input" placeholder="false 提示词...">
                </div>
            `;
            aiSettingsDiv.appendChild(booleanPromptsContainer);
            aiSubContainer.appendChild(aiSettingsDiv);
            bodyDiv.appendChild(aiSubContainer);

            function renderManualSubUI() {
                manualOptionsContainer.innerHTML = '';
                if (cfg.inputType !== 'manual') return;

                if (cfg.manualSubType === 'boolean') {
                    const booleanDiv = document.createElement('div');
                    booleanDiv.className = 'boolean-fields';
                    const actionsRow = document.createElement('div');
                    actionsRow.className = 'boolean-actions-row';

                    function createActionItem(labelText, actionKey) {
                        const item = document.createElement('div');
                        item.className = 'boolean-action-item';
                        const lbl = document.createElement('label');
                        lbl.textContent = labelText;
                        const selectEl = document.createElement('select');
                        selectEl.className = `custom-select-mini ${actionKey}-action-select`;
                        selectEl.innerHTML = '<option value="none">无</option><option value="fixed">固定值</option><option value="ai">接入AI</option>';
                        selectEl.value = cfg[actionKey];
                        const fixedInput = document.createElement('input');
                        fixedInput.type = 'text';
                        fixedInput.className = 'fixed-value-input';
                        fixedInput.placeholder = '固定值文本';
                        fixedInput.value = cfg[actionKey + 'FixedText'];
                        fixedInput.style.display = (selectEl.value === 'fixed') ? 'block' : 'none';
                        selectEl.addEventListener('change', () => {
                            cfg[actionKey] = selectEl.value;
                            fixedInput.style.display = (selectEl.value === 'fixed') ? 'block' : 'none';
                            updateBooleanPrompts();
                        });
                        fixedInput.addEventListener('input', () => { cfg[actionKey + 'FixedText'] = fixedInput.value; });
                        item.appendChild(lbl);
                        item.appendChild(selectEl);
                        item.appendChild(fixedInput);
                        return item;
                    }
                    actionsRow.appendChild(createActionItem('false 执行', 'falseAction'));
                    actionsRow.appendChild(createActionItem('true 执行', 'trueAction'));
                    booleanDiv.appendChild(actionsRow);
                    const defaultGroup = document.createElement('div');
                    defaultGroup.className = 'input-group';
                    const defaultLabel = document.createElement('label');
                    defaultLabel.textContent = '默认值';
                    const defaultSelect = document.createElement('select');
                    defaultSelect.className = 'custom-select-mini';
                    defaultSelect.innerHTML = '<option value="none">无</option><option value="true">true</option><option value="false">false</option>';
                    defaultSelect.value = cfg.booleanDefault;
                    defaultSelect.addEventListener('change', () => { cfg.booleanDefault = defaultSelect.value; });
                    defaultGroup.appendChild(defaultLabel);
                    defaultGroup.appendChild(defaultSelect);
                    booleanDiv.appendChild(defaultGroup);
                    manualOptionsContainer.appendChild(booleanDiv);

                    const truePromptInput = booleanPromptsContainer.querySelector('.true-prompt-input');
                    const falsePromptInput = booleanPromptsContainer.querySelector('.false-prompt-input');
                    truePromptInput.value = cfg.truePrompt || '';
                    falsePromptInput.value = cfg.falsePrompt || '';
                    truePromptInput.addEventListener('input', () => { cfg.truePrompt = truePromptInput.value; });
                    falsePromptInput.addEventListener('input', () => { cfg.falsePrompt = falsePromptInput.value; });

                    function updateBooleanPrompts() {
                        const showTrue = cfg.trueAction === 'ai';
                        const showFalse = cfg.falseAction === 'ai';
                        const anyAI = showTrue || showFalse;
                        if (anyAI) {
                            aiCheckbox.checked = true;
                            cfg.aiEnabled = true;
                            aiSettingsDiv.style.display = 'block';
                            aiCheckbox.disabled = true;
                        } else {
                            aiCheckbox.disabled = false;
                        }
                        generalPromptGroup.style.display = anyAI ? 'none' : '';
                        booleanPromptsContainer.style.display = anyAI ? 'block' : 'none';
                        booleanPromptsContainer.querySelector('[data-prompt-type="true"]').style.display = showTrue ? '' : 'none';
                        booleanPromptsContainer.querySelector('[data-prompt-type="false"]').style.display = showFalse ? '' : 'none';
                    }
                    updateBooleanPrompts();
                } else if (cfg.manualSubType === 'choice') {
                    const choiceDiv = document.createElement('div');
                    choiceDiv.className = 'choice-fields';
                    const currentChoiceType = cfg.choiceType;
                    if (currentChoiceType === 'multiple') { if (!cfg.choiceDefaultIndexes) cfg.choiceDefaultIndexes = []; }
                    else { if (cfg.choiceDefaultIndex === undefined) cfg.choiceDefaultIndex = 0; }
                    const optionsContainer = document.createElement('div');
                    optionsContainer.className = 'choice-options-list';
                    function renderOptions() {
                        optionsContainer.innerHTML = '';
                        cfg.choiceOptions.forEach((opt, idx) => {
                            const row = document.createElement('div');
                            row.className = 'choice-option-row';
                            let selectInput;
                            if (currentChoiceType === 'multiple') {
                                selectInput = document.createElement('input');
                                selectInput.type = 'checkbox';
                                selectInput.checked = cfg.choiceDefaultIndexes.includes(idx);
                                selectInput.addEventListener('change', () => {
                                    if (selectInput.checked) { if (!cfg.choiceDefaultIndexes.includes(idx)) cfg.choiceDefaultIndexes.push(idx); }
                                    else { cfg.choiceDefaultIndexes = cfg.choiceDefaultIndexes.filter(i => i !== idx); }
                                });
                            } else {
                                selectInput = document.createElement('input');
                                selectInput.type = 'radio';
                                selectInput.name = `choice-default-${tag}`;
                                selectInput.checked = (cfg.choiceDefaultIndex === idx);
                                selectInput.addEventListener('change', () => { cfg.choiceDefaultIndex = idx; });
                            }
                            const input = document.createElement('input');
                            input.type = 'text';
                            input.value = opt;
                            input.placeholder = `选项 ${idx + 1}`;
                            input.style.flex = '1';
                            input.style.maxWidth = '200px';
                            input.addEventListener('input', (e) => { cfg.choiceOptions[idx] = e.target.value; });
                            const deleteBtn = document.createElement('button');
                            deleteBtn.className = 'icon-btn small';
                            deleteBtn.innerHTML = `<svg viewBox="0 0 24 24" width="12" height="12"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>`;
                            deleteBtn.style.color = '#FF3B30';
                            deleteBtn.style.marginLeft = '4px';
                            deleteBtn.onclick = () => {
                                cfg.choiceOptions.splice(idx, 1);
                                if (currentChoiceType === 'multiple') {
                                    cfg.choiceDefaultIndexes = cfg.choiceDefaultIndexes.filter(i => i !== idx).map(i => i > idx ? i - 1 : i);
                                } else {
                                    if (cfg.choiceDefaultIndex >= cfg.choiceOptions.length) cfg.choiceDefaultIndex = cfg.choiceOptions.length - 1;
                                    else if (cfg.choiceDefaultIndex > idx) cfg.choiceDefaultIndex--;
                                }
                                renderOptions();
                            };
                            if (cfg.choiceOptions.length <= 2) { deleteBtn.disabled = true; deleteBtn.style.opacity = '0.3'; }
                            row.appendChild(selectInput);
                            row.appendChild(input);
                            row.appendChild(deleteBtn);
                            optionsContainer.appendChild(row);
                        });
                        const addBtn = document.createElement('button');
                        addBtn.className = 'add-option-btn';
                        addBtn.textContent = '+ 添加选项';
                        addBtn.onclick = () => { cfg.choiceOptions.push('新选项'); renderOptions(); };
                        optionsContainer.appendChild(addBtn);
                    }
                    renderOptions();
                    choiceDiv.appendChild(optionsContainer);
                    manualOptionsContainer.appendChild(choiceDiv);
                } else if (cfg.manualSubType === 'text') {
                    const textGroup = document.createElement('div');
                    textGroup.className = 'input-group';
                    const displayLabel = document.createElement('label');
                    displayLabel.textContent = '显示文本';
                    const displayInput = document.createElement('input');
                    displayInput.type = 'text';
                    displayInput.className = 'display-text-input';
                    displayInput.placeholder = '输入框中的灰色提示文字';
                    displayInput.value = cfg.displayText || '';
                    displayInput.addEventListener('input', () => { cfg.displayText = displayInput.value; });
                    textGroup.appendChild(displayLabel);
                    textGroup.appendChild(displayInput);
                    manualOptionsContainer.appendChild(textGroup);
                } else if (cfg.manualSubType === 'image') {
                    const hint = document.createElement('div');
                    hint.style.fontSize = '12px';
                    hint.style.color = 'var(--text-sub)';
                    hint.textContent = '支持剪贴板导入和本地添加，可填写题注和简介';
                    manualOptionsContainer.appendChild(hint);

                    const captionSwitchContainer = document.createElement('div');
                    captionSwitchContainer.style.display = 'flex';
                    captionSwitchContainer.style.alignItems = 'center';
                    captionSwitchContainer.style.gap = '4px';
                    captionSwitchContainer.style.marginTop = '8px';
                    const captionLabel = document.createElement('label');
                    captionLabel.textContent = '生成图注';
                    const captionSwitchEl = document.createElement('label');
                    captionSwitchEl.className = 'switch';
                    const captionCheckbox = document.createElement('input');
                    captionCheckbox.type = 'checkbox';
                    captionCheckbox.checked = cfg.generateCaption;
                    captionCheckbox.addEventListener('change', () => { cfg.generateCaption = captionCheckbox.checked; });
                    const captionSlider = document.createElement('span');
                    captionSlider.className = 'slider round';
                    captionSwitchEl.appendChild(captionCheckbox);
                    captionSwitchEl.appendChild(captionSlider);
                    captionSwitchContainer.appendChild(captionLabel);
                    captionSwitchContainer.appendChild(captionSwitchEl);
                    manualOptionsContainer.appendChild(captionSwitchContainer);
                }
            }

            renderManualSubUI();

            inputTypeSelect.addEventListener('change', () => {
                cfg.inputType = inputTypeSelect.value;
                const showManual = cfg.inputType === 'manual';
                typeLabel.style.display = showManual ? '' : 'none';
                manualTypeSelect.style.display = showManual ? '' : 'none';
                const showChoiceType = showManual && manualTypeSelect.value === 'choice';
                choiceTypeLabel.style.display = showChoiceType ? '' : 'none';
                choiceTypeSelect.style.display = showChoiceType ? '' : 'none';
                autoInputGroup.style.display = (cfg.inputType === 'none' && !cfg.aiEnabled) ? '' : 'none';
                renderManualSubUI();
            });
            manualTypeSelect.addEventListener('change', () => {
                cfg.manualSubType = manualTypeSelect.value;
                const showChoiceType = cfg.inputType === 'manual' && manualTypeSelect.value === 'choice';
                choiceTypeLabel.style.display = showChoiceType ? '' : 'none';
                choiceTypeSelect.style.display = showChoiceType ? '' : 'none';
                renderManualSubUI();
            });
            choiceTypeSelect.addEventListener('change', () => {
                cfg.choiceType = choiceTypeSelect.value;
                if (cfg.choiceType === 'multiple') cfg.choiceDefaultIndexes = [];
                else cfg.choiceDefaultIndex = 0;
                renderManualSubUI();
            });
            aiCheckbox.addEventListener('change', () => {
                cfg.aiEnabled = aiCheckbox.checked;
                aiSettingsDiv.style.display = cfg.aiEnabled ? 'block' : 'none';
                autoInputGroup.style.display = (cfg.inputType === 'none' && !cfg.aiEnabled) ? '' : 'none';
            });
            thinkCheckbox.addEventListener('change', () => { cfg.thinkingMode = thinkCheckbox.checked; });

            itemDiv.appendChild(headerDiv);
            itemDiv.appendChild(bodyDiv);
            listContainer.appendChild(itemDiv);
        });
    } catch (error) {
        console.error("Word 抓取控件异常:", error);
        emptyState.style.display = 'block';
        listContainer.style.display = 'none';
    } finally {
        isWordRefreshing = false;
    }
}

async function writeWordCustomData(template, aiData) {
    return Word.run(async (context) => {
        const contentControls = context.document.contentControls;
        contentControls.load("items/tag, items/title");
        await context.sync();
        const config = template.controlsConfig || {};
        let writeCount = 0;

        for (const [tag, cfg] of Object.entries(config)) {
            if (cfg.manualSubType === 'boolean' && cfg.inputType === 'manual') {
                const booleanValue = aiData[tag];
                const realValue = (booleanValue === 'true');
                const falseAction = cfg.falseAction;
                const trueAction = cfg.trueAction;
                let finalText = '';
                if (realValue && trueAction === 'fixed') finalText = cfg.trueActionFixedText || '';
                else if (realValue && trueAction === 'ai') finalText = aiData[tag] || '';
                else if (!realValue && falseAction === 'fixed') finalText = cfg.falseActionFixedText || '';
                else if (!realValue && falseAction === 'ai') finalText = aiData[tag] || '';
                if (finalText.trim()) {
                    const targetControls = contentControls.items.filter(c => c.tag === tag || c.title === tag);
                    for (const control of targetControls) {
                        control.insertText(finalText, "Replace");
                        control.font.size = 12;
                        control.font.name = "宋体";
                        writeCount++;
                    }
                }
                continue;
            }

            if (cfg.manualSubType === 'image' && cfg.inputType === 'manual') {
                const images = cfg.images || [];
                const hasAI = cfg.aiEnabled;
                const prompt = cfg.aiPrompt;
                const genCaption = cfg.generateCaption;
                const targetControls = contentControls.items.filter(c => c.tag === tag || c.title === tag);
                for (const control of targetControls) {
                    if (!hasAI) {
                        for (const img of images) {
                            const pureBase64 = img.base64.split(',')[1];
                            control.insertInlinePictureFromBase64(pureBase64, "End");
                        }
                    } else if (hasAI && !prompt && genCaption) {
                        const captionText = aiData[tag] || '';
                        for (const img of images) {
                            const pureBase64 = img.base64.split(',')[1];
                            control.insertInlinePictureFromBase64(pureBase64, "End");
                            if (captionText) control.insertText(`\n图注：${captionText}`, "End");
                        }
                    } else if (hasAI && prompt) {
                        const text = aiData[tag] || '';
                        if (text.trim()) control.insertText(text, "Replace");
                    }
                    writeCount++;
                }
                continue;
            }

            let finalText = '';
            if (cfg.aiEnabled && aiData[tag] && aiData[tag].trim()) finalText = aiData[tag];
            else if (cfg.autoInputText && cfg.autoInputText.trim()) finalText = cfg.autoInputText;
            else if (aiData[tag] && aiData[tag].trim()) finalText = aiData[tag];

            if (!finalText.trim()) continue;

            const targetControls = contentControls.items.filter(c => c.tag === tag || c.title === tag);
            for (const control of targetControls) {
                control.insertText(finalText, "Replace");
                control.font.size = 12;
                control.font.name = "宋体";
                writeCount++;
            }
        }
        await context.sync();
        if (writeCount === 0) {
            const body = context.document.body;
            let insertText = '';
            for (const [tag, cfg] of Object.entries(config)) {
                let finalText = '';
                if (cfg.aiEnabled && aiData[tag] && aiData[tag].trim()) finalText = aiData[tag];
                else if (cfg.autoInputText && cfg.autoInputText.trim()) finalText = cfg.autoInputText;
                else if (aiData[tag] && aiData[tag].trim()) finalText = aiData[tag];
                if (finalText.trim()) insertText += `${tag}：\n${finalText}\n\n`;
            }
            if (insertText.trim()) body.insertText(insertText.trim(), Word.InsertLocation.end);
        }
        await context.sync();
    });
}