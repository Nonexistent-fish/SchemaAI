// ========== 自定义模板相关代码 ==========
let controlsConfig = {};
let isRefreshing = false;

Office.onReady(() => {
    document.getElementById("save-template-btn").addEventListener("click", () => {
        const titleInput = document.getElementById("custom-template-title");
        const title = titleInput.value.trim();
        const desc = document.getElementById("custom-template-desc").value.trim();

        document.querySelectorAll('.error-flash').forEach(el => el.classList.remove('error-flash'));

        let hasError = false;

        if (!title) {
            titleInput.closest('.input-group').classList.add('error-flash');
            hasError = true;
            setTimeout(() => titleInput.closest('.input-group').classList.remove('error-flash'), 400);
        }

        document.querySelectorAll('.control-item').forEach(item => {
            const ctrlTitleInput = item.querySelector('.control-title-input');
            if (ctrlTitleInput && !ctrlTitleInput.value.trim()) {
                ctrlTitleInput.closest('.control-header-left').classList.add('error-flash');
                hasError = true;
                setTimeout(() => ctrlTitleInput.closest('.control-header-left').classList.remove('error-flash'), 400);
            }
        });

        if (hasError) return;

        CONFIG.customTemplateTitle = title;
        CONFIG.customTemplateDesc = desc;
        const cleanConfig = {};
        const titles = {};
        Object.keys(controlsConfig).forEach(tag => {
            cleanConfig[tag] = sanitizeConfig(controlsConfig[tag]);
            const itemDiv = document.querySelector(`.control-item[data-tag="${tag}"]`);
            if (itemDiv) {
                const ti = itemDiv.querySelector('.control-title-input');
                titles[tag] = ti ? ti.value.trim() : '';
            }
        });
        CONFIG.customTemplateData = { title, desc, controlsConfig: cleanConfig, titles };
        saveStorage();
        showStatus("模板已保存", "success");
        switchView("view-home");
    });

    document.getElementById("edit-template-btn").addEventListener("click", () => {
        switchView("view-custom");
        document.getElementById("custom-template-title").value = CONFIG.customTemplateTitle || "";
        document.getElementById("custom-template-desc").value = CONFIG.customTemplateDesc || "";
        if (CONFIG.customTemplateData && CONFIG.customTemplateData.controlsConfig) {
            const saved = CONFIG.customTemplateData.controlsConfig;
            controlsConfig = {};
            Object.keys(saved).forEach(tag => { controlsConfig[tag] = sanitizeConfig(saved[tag]); });
        } else {
            controlsConfig = {};
        }
        refreshCustomView();
    });

    document.getElementById("refresh-custom-btn").addEventListener("click", refreshCustomView);
});

function sanitizeConfig(cfg) {
    const def = {
        inputType: 'none', manualSubType: 'boolean', aiEnabled: false, aiPrompt: '',
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
                clean[key] = value.map(item => typeof item === 'string' ? sanitizeString(item) : item);
                if (key === 'choiceOptions' && clean[key].length < 2) clean[key] = def[key];
            }
        } else if (typeof def[key] === 'boolean') {
            clean[key] = !!value;
        } else if (typeof def[key] === 'number') {
            clean[key] = isNaN(Number(value)) ? def[key] : Number(value);
        } else {
            clean[key] = sanitizeString(value);
        }
    });
    return clean;
}

function sanitizeString(str) {
    if (typeof str !== 'string') return '';
    return str.replace(/[\uFFFD\u0000-\u001F\u007F-\u009F]/g, '').trim();
}

async function refreshCustomView() {
    if (isRefreshing) return;
    isRefreshing = true;
    const listContainer = document.getElementById("custom-controls-list");
    const emptyState = document.getElementById("custom-empty-state");
    listContainer.innerHTML = '';
    listContainer.style.display = 'none';
    emptyState.style.display = 'none';
    try {
        await Word.run(async (context) => {
            const contentControls = context.document.contentControls;
            contentControls.load('items');
            await context.sync();
            const controlsWithTag = contentControls.items.filter(c => c.tag && c.tag.trim() !== '');
            if (controlsWithTag.length === 0) { emptyState.style.display = 'block'; return; }
            listContainer.style.display = 'block';
            const groups = {};
            controlsWithTag.forEach(c => {
                const tag = c.tag.trim();
                if (!groups[tag]) groups[tag] = { tag, title: c.title || '', controls: [] };
                groups[tag].controls.push(c);
            });
            Object.keys(groups).forEach(tag => {
                const group = groups[tag];
                controlsConfig[tag] = sanitizeConfig(controlsConfig[tag] || {});
                const cfg = controlsConfig[tag];

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

                // AI 设置区域
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

                // AI 提示词
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
        });
    } catch (error) { emptyState.style.display = 'block'; listContainer.style.display = 'none'; }
    finally { isRefreshing = false; }
}

function initCustomView() {
    if (CONFIG.customTemplateData) {
        const cfg = CONFIG.customTemplateData.controlsConfig;
        if (!cfg || typeof cfg !== 'object' || Object.keys(cfg).length === 0) {
            CONFIG.customTemplateData = null;
            saveStorage();
        }
    }
    refreshCustomView();
}

function renderTemplateUseView() {
    const template = CONFIG.customTemplateData;
    if (!template) return;
    document.getElementById("template-use-title").innerText = template.title || "模板";
    document.getElementById("template-use-desc").innerText = template.desc || "";
    const container = document.getElementById("template-use-controls-list");
    container.innerHTML = "";
    const config = template.controlsConfig || {};
    const titles = template.titles || {};

    Object.keys(config).forEach(tag => {
        const cfg = sanitizeConfig(config[tag]);
        if (cfg.inputType === 'none' && !cfg.autoInputText) return;
        if (cfg.inputType === 'none' && cfg.autoInputText) return;

        const itemDiv = document.createElement('div');
        itemDiv.className = 'template-use-item';
        const headerDiv = document.createElement('div');
        headerDiv.className = 'control-header';
        headerDiv.style.cursor = 'default';
        const displayTitle = titles[tag] || tag;
        headerDiv.innerHTML = `
            <div class="control-header-left">
                <span style="font-weight:500;font-size:14px;">${displayTitle}</span>
                <span class="control-tag">${tag}</span>
            </div>
        `;
        itemDiv.appendChild(headerDiv);

        if (cfg.inputType === 'manual') {
            if (cfg.manualSubType === 'text') {
                const input = document.createElement('input');
                input.type = 'text';
                input.className = 'template-use-input';
                input.disabled = false;
                input.placeholder = cfg.displayText || '请输入文本';
                itemDiv.appendChild(input);
            } else if (cfg.manualSubType === 'boolean') {
                const label = document.createElement('label');
                label.className = 'switch';
                const checkbox = document.createElement('input');
                checkbox.type = 'checkbox';
                checkbox.disabled = false;
                checkbox.checked = (cfg.booleanDefault === 'true');
                label.appendChild(checkbox);
                const slider = document.createElement('span');
                slider.className = 'slider round';
                label.appendChild(slider);
                itemDiv.appendChild(label);
            } else if (cfg.manualSubType === 'choice') {
                const options = cfg.choiceOptions;
                if (cfg.choiceType === 'dropdown') {
                    const selectEl = document.createElement('select');
                    selectEl.className = 'template-use-input';
                    selectEl.disabled = false;
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
                        const input = document.createElement('input');
                        if (cfg.choiceType === 'multiple') {
                            input.type = 'checkbox';
                            input.checked = (cfg.choiceDefaultIndexes || []).includes(idx);
                        } else {
                            input.type = 'radio';
                            input.name = `use-choice-${tag}`;
                            input.checked = (cfg.choiceDefaultIndex === idx);
                        }
                        input.disabled = false;
                        const text = document.createElement('span');
                        text.textContent = opt;
                        text.style.fontSize = '13px';
                        row.appendChild(input);
                        row.appendChild(text);
                        itemDiv.appendChild(row);
                    });
                }
            } else if (cfg.manualSubType === 'image') {
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
                        imgDiv.onclick = () => {
                            document.getElementById("current-img-id").value = idx;
                            document.getElementById("current-img-desc").value = img.desc || '';
                            document.getElementById("modal-overlay").classList.add("open");
                            document.getElementById("img-desc-modal").classList.add("open");
                            window._currentImageControl = cfg;
                        };
                        galleryDiv.appendChild(imgDiv);
                    });
                };
                renderGallery();

                const buttonRow = document.createElement('div');
                buttonRow.style.display = 'flex';
                buttonRow.style.gap = '8px';
                buttonRow.style.marginTop = '8px';

                const clipboardBtn = document.createElement('button');
                clipboardBtn.className = 'icon-btn small';
                clipboardBtn.innerHTML = `<svg viewBox="0 0 24 24"><path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2"></path><rect x="8" y="2" width="8" height="4" rx="1" ry="1"></rect></svg>`;
                clipboardBtn.title = '从剪贴板粘贴';
                clipboardBtn.onclick = async () => {
                    try {
                        const items = await navigator.clipboard.read();
                        for (const item of items) {
                            const imageTypes = item.types.filter(t => t.startsWith('image/'));
                            if (imageTypes.length > 0) {
                                const blob = await item.getType(imageTypes[0]);
                                const reader = new FileReader();
                                reader.onload = (e) => {
                                    cfg.images.push({ base64: e.target.result, caption: '', desc: '' });
                                    renderGallery();
                                };
                                reader.readAsDataURL(blob);
                                return;
                            }
                        }
                    } catch (err) {
                        showStatus('请授予剪贴板权限，或直接按 Ctrl+V 粘贴', 'warn');
                    }
                };

                const addBtn = document.createElement('button');
                addBtn.className = 'icon-btn small';
                addBtn.innerHTML = `<svg viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"></line><line x1="5" y1="12" x2="19" y2="12"></line></svg>`;
                addBtn.title = '导入本地图片';
                addBtn.onclick = () => {
                    const input = document.createElement('input');
                    input.type = 'file';
                    input.accept = 'image/*';
                    input.onchange = (e) => {
                        const files = Array.from(e.target.files);
                        files.forEach(file => {
                            const reader = new FileReader();
                            reader.onload = (event) => {
                                cfg.images.push({ base64: event.target.result, caption: '', desc: '' });
                                renderGallery();
                            };
                            reader.readAsDataURL(file);
                        });
                    };
                    input.click();
                };

                buttonRow.appendChild(clipboardBtn);
                buttonRow.appendChild(addBtn);
                itemDiv.appendChild(galleryDiv);
                itemDiv.appendChild(buttonRow);
            }
        }
        container.appendChild(itemDiv);
    });

    const bar = document.createElement("div");
    bar.className = "generate-action-bar";
    bar.innerHTML = `
        <button id="template-generate-btn" class="primary-btn" style="margin-top:0;">执行生成</button>
        <button id="template-save-doc-btn" class="save-icon-btn" title="一键保存文档">
            <svg viewBox="0 0 24 24" width="20" height="20" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
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

    document.getElementById("template-generate-btn").addEventListener("click", async () => {
        const btn = document.getElementById("template-generate-btn");
        const saveBtn = document.getElementById("template-save-doc-btn");
        const progressBar = progressContainer.querySelector('.progress-bar');
        const updateProgress = (width) => { progressContainer.style.display = 'block'; progressBar.style.width = width + '%'; };
        const showStatus = (text, type = 'info') => { statusMsg.innerHTML = `<span class="status-${type}">${text}</span>`; };
        showStatus('', 'info');
        saveBtn.classList.remove("show");

        if (!CONFIG.url) { showStatus("请先配置API接口", "error"); return; }

        const modelNameDisplay = CONFIG.model || "本地大模型";
        const userInputs = collectUserInputs();

        try {
            btn.disabled = true; btn.innerText = "生成中..."; updateProgress(10);
            showStatus(`正在请求 ${modelNameDisplay}...`, "info");
            const aiData = await fetchCustomAIResponse(template, userInputs);
            updateProgress(70);
            showStatus(`等待 ${modelNameDisplay} 响应...`, "info");
            await writeCustomDataToWord(template, aiData);
            updateProgress(100);
            showStatus("生成成功", "success");
            setTimeout(() => { progressContainer.style.display = "none"; saveBtn.classList.add("show"); }, 1000);
        } catch (error) {
            progressBar.style.background = "var(--error-color)";
            showStatus(`错误：${error.message}`, "error");
        } finally {
            btn.disabled = false;
            setTimeout(() => { btn.innerText = "执行生成"; progressBar.style.background = ""; }, 2500);
        }
    });
    document.getElementById("template-save-doc-btn").addEventListener("click", handleSaveDocument);
}

function collectUserInputs() {
    const inputs = {};
    const items = document.querySelectorAll('.template-use-item');
    items.forEach(item => {
        const tagEl = item.querySelector('.control-tag');
        if (!tagEl) return;
        const tag = tagEl.textContent.trim();
        const textInput = item.querySelector('input[type="text"].template-use-input');
        if (textInput) { inputs[tag] = textInput.value.trim(); return; }
        const checkbox = item.querySelector('input[type="checkbox"]');
        if (checkbox && item.querySelector('.switch')) { inputs[tag] = checkbox.checked ? 'true' : 'false'; return; }
        const radios = item.querySelectorAll('input[type="radio"]');
        if (radios.length > 0) {
            const checkedRadio = item.querySelector('input[type="radio"]:checked');
            if (checkedRadio) {
                const parentRow = checkedRadio.closest('.choice-option-row');
                const span = parentRow ? parentRow.querySelector('span') : null;
                inputs[tag] = span ? span.textContent.trim() : '';
            }
            return;
        }
        const checkboxes = item.querySelectorAll('input[type="checkbox"]:not(.switch input)');
        if (checkboxes.length > 0) {
            const checkedBoxes = Array.from(checkboxes).filter(cb => cb.checked);
            const values = [];
            checkedBoxes.forEach(cb => {
                const parentRow = cb.closest('.choice-option-row');
                const span = parentRow ? parentRow.querySelector('span') : null;
                if (span) values.push(span.textContent.trim());
            });
            inputs[tag] = values.join(', ');
            return;
        }
        const selectEl = item.querySelector('select');
        if (selectEl) { inputs[tag] = selectEl.value !== '' ? selectEl.options[selectEl.selectedIndex].text : ''; return; }
        const imgItems = item.querySelectorAll('.image-item');
        if (imgItems.length > 0) {
            const descs = [];
            imgItems.forEach(imgDiv => {
                const descEl = imgDiv.querySelector('.img-preview');
                if (descEl && descEl.title) descs.push(descEl.title.replace('点击添加图片描述', '').trim());
            });
            inputs[tag] = descs.filter(d => d).join('; ');
        }
    });
    return inputs;
}

async function fetchCustomAIResponse(template, userInputs = {}) {
    const config = template.controlsConfig || {};
    const titles = template.titles || {};
    const tags = Object.keys(config);
    let systemPrompt = `你是文档内容生成助手。请根据以下信息生成内容，并返回一个纯JSON对象。\n`;
    systemPrompt += `模板标题：${template.title}\n模板描述：${template.desc}\n\n字段列表：\n`;
    tags.forEach(tag => {
        const cfg = config[tag];
        const title = titles[tag] || tag;
        systemPrompt += `- 标记："${tag}"，标题："${title}"`;
        let extra = '';
        if (cfg.aiEnabled && cfg.aiPrompt) {
            extra = `，额外要求：${cfg.aiPrompt}`;
        }
        if (cfg.manualSubType === 'boolean' && userInputs[tag] !== undefined) {
            const boolVal = userInputs[tag];
            if (boolVal === 'true' && cfg.trueAction === 'ai' && cfg.truePrompt) {
                extra = `，额外要求：${cfg.truePrompt}`;
            } else if (boolVal === 'false' && cfg.falseAction === 'ai' && cfg.falsePrompt) {
                extra = `，额外要求：${cfg.falsePrompt}`;
            }
        }
        systemPrompt += extra;
        if (userInputs[tag] !== undefined && userInputs[tag] !== '') {
            systemPrompt += `，当前用户输入：${userInputs[tag]}`;
        }
        systemPrompt += `\n`;
    });
    systemPrompt += `\n重要规则：\n1. 只返回一个JSON对象，其键必须严格等于上述标记（tag），不能有任何其他字段。\n2. 每个键对应的值是你为那个字段生成的内容。\n3. 不要包含任何解释、说明、Markdown或多余文字。\n示例返回：{"number_1":"内容1","number_2":"内容2","mine":"内容3"}`;

    const messagesPayload = [{ role: "user", content: systemPrompt }];
    const isDeepSeekV4 = /^deepseek(-v4)?-(flash|pro)$/i.test(CONFIG.model);
    const requestBody = { model: CONFIG.model, messages: messagesPayload, max_tokens: 4096 };
    const anyThinking = Object.values(config).some(c => c.thinkingMode === true);
    if (isDeepSeekV4 && anyThinking) {
        requestBody.thinking = { type: "enabled" };
    } else if (!isDeepSeekV4) {
        requestBody.temperature = 0.7;
    }

    const headers = { "Content-Type": "application/json" };
    if (CONFIG.key) headers["Authorization"] = `Bearer ${CONFIG.key}`;
    const response = await fetch(CONFIG.url, { method: "POST", headers, body: JSON.stringify(requestBody) });
    if (!response.ok) { const errText = await response.text().catch(() => "无返回"); throw new Error(`AI 请求失败 (${response.status}): ${errText}`); }
    let rawText = (await response.json()).choices[0].message.content;
    rawText = rawText.replace(/<think>[\s\S]*?<\/think>/g, '').trim().replace(/```json|```/gi, '').trim();
    try { return JSON.parse(rawText); } catch (e) {
        const fixed = rawText.replace(/,\s*\}/g, '}');
        try { return JSON.parse(fixed); } catch (e2) { throw new Error("AI 返回的数据格式错误，请重试"); }
    }
}

async function writeCustomDataToWord(template, aiData) {
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
                            if (captionText) {
                                control.insertText(`\n图注：${captionText}`, "End");
                            }
                        }
                    } else if (hasAI && prompt) {
                        const text = aiData[tag] || '';
                        if (text.trim()) {
                            control.insertText(text, "Replace");
                        }
                    }
                    writeCount++;
                }
                continue;
            }

            let finalText = '';
            if (cfg.aiEnabled && aiData[tag] && aiData[tag].trim()) {
                finalText = aiData[tag];
            } else if (cfg.autoInputText && cfg.autoInputText.trim()) {
                finalText = cfg.autoInputText;
            } else if (aiData[tag] && aiData[tag].trim()) {
                finalText = aiData[tag];
            }
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
    }).catch(e => {
        if (e instanceof OfficeExtension.Error && e.code === "AccessDenied") throw new Error("文档只读，无法写入");
        throw e;
    });
}