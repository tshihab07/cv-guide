function addSkillCategoryRow() {
    const container = document.getElementById('skill-category-rows');
    let template;
    
    if (container.children.length > 0) {
        template = container.children[0].cloneNode(true);
        template.querySelectorAll('input, textarea').forEach(i => {
            i.value = '';
        });
        template.querySelector('.skill-custom-category').value = '';
    }
    else {
        template = document.createElement('div');
        template.className = 'skill-category-row form-grid';
        template.style.cssText = 'margin-bottom: 16px; padding: 16px; border: 1px solid var(--border); border-radius: 12px; background: var(--surface);';
        
        let optionsHtml = '';
        defaultSkillCategories.forEach(cat => {
            optionsHtml += `<li class="combobox-option" data-value="${cat}">${cat}</li>`;
        });
        
        optionsHtml += `<li class="combobox-option combobox-custom" data-value="custom">+ Add custom category...</li>`;
        template.innerHTML = `
            <label>Category
                <div class="combobox-wrapper">
                    <input type="text" name="skill_category[]" class="form-control combobox-input skill-category-combobox" autocomplete="off" placeholder="Select or type category" aria-expanded="false" aria-controls="combobox-options" />
                    <input type="hidden" name="skill_custom_category[]" class="skill-custom-category" value="" />
                    <div class="combobox-dropdown" role="listbox">
                        <ul class="combobox-options" id="combobox-options" role="presentation">
                            ${optionsHtml}
                        </ul>
                    </div>
                </div>
            </label>
            <label>Skills (comma separated)<textarea name="skill_items[]" rows="2"></textarea></label>
            <button type="button" class="btn btn-danger btn-sm delete-row-btn" style="align-self: flex-start; height: fit-content; padding: 8px 12px;" onclick="deleteRow(this)">Delete</button>
        `;
    }

    container.appendChild(template);
    initCombobox(template.querySelector('.skill-category-combobox'));
}

function addProjectRow() {
    const container = document.getElementById('project-rows');
    let template;
    
    if (container.children.length > 0) {
        template = container.children[0].cloneNode(true);
        template.querySelectorAll('input, textarea').forEach(i => { i.value = ''; });
        
        // reset combobox if present
        const combobox = template.querySelector('.project-category-combobox');
        if (combobox) {
            combobox.value = '';
        }

        // add delete button
        if (!template.querySelector('.delete-row-btn')) {
            const deleteBtn = document.createElement('button');
            deleteBtn.type = 'button';
            deleteBtn.className = 'btn btn-danger btn-sm delete-row-btn';
            deleteBtn.style.cssText = 'margin-top: 12px; padding: 8px 12px;';
            deleteBtn.textContent = 'Delete';
            deleteBtn.onclick = function() { deleteRow(this); };
            template.appendChild(deleteBtn);
        }
    }
    else {
        template = document.createElement('div');
        template.className = 'project-row';
        template.style.cssText = 'margin-bottom: 20px; padding: 16px; border: 1px solid var(--border); border-radius: 12px; background: var(--surface);';
        
        let optionsHtml = '';
        defaultProjectCategories.forEach(cat => {
            optionsHtml += `<li class="combobox-option" data-value="${cat}">${cat}</li>`;
        });
        
        optionsHtml += `<li class="combobox-option combobox-custom" data-value="custom">+ Add custom category...</li>`;
        template.innerHTML = `
            <div class="form-grid">
                <label>Project Title<input name="project_title[]" placeholder="Intelligent Loan Approval Prediction System" /></label>
                <label>Project Category
                    <div class="combobox-wrapper">
                        <input type="text" name="project_category[]" class="form-control combobox-input project-category-combobox" autocomplete="off" placeholder="Select or type category" aria-expanded="false" />
                        <div class="combobox-dropdown" role="listbox">
                            <ul class="combobox-options" role="presentation">
                                ${optionsHtml}
                            </ul>
                        </div>
                    </div>
                </label>
                <label>Project Description<textarea name="project_description[]" rows="4" placeholder="Performed EDA and feature engineering on financial datasets...&#10;Reduced predicted default risk by 86% using XGBoost with optuna optimization&#10;Ensure fair and consistent evaluation&#10;Cut manual screening time by 67% through automated ML pipeline"></textarea></label>
                <label>Source Code Link (optional)<input name="source_code_link[]" placeholder="https://github.com/..." /></label>
                <label>Project Link (optional)<input name="project_link[]" placeholder="https://your-deployed-url.com/" /></label>
            </div>
            <button type="button" class="btn btn-danger btn-sm delete-row-btn" style="margin-top: 12px; padding: 8px 12px;" onclick="deleteRow(this)">Delete</button>
        `;
    }

    container.appendChild(template);
    
    // initialize combobox for new row
    const combobox = template.querySelector('.project-category-combobox');
    if (combobox) {
        initCombobox(combobox);
    }
}

function addExperienceRow() {
    const container = document.getElementById('experience-rows');
    let template;
    
    if (container.children.length > 0) {
        template = container.children[0].cloneNode(true);
        template.querySelectorAll('input, textarea').forEach(i => { i.value = ''; });
        
        // add delete button if not present
        if (!template.querySelector('.delete-row-btn')) {
            const deleteBtn = document.createElement('button');
            deleteBtn.type = 'button';
            deleteBtn.className = 'btn btn-danger btn-sm delete-row-btn';
            deleteBtn.style.cssText = 'margin-top: 12px; padding: 8px 12px;';
            deleteBtn.textContent = 'Delete';
            deleteBtn.onclick = function() { deleteRow(this); };
            template.appendChild(deleteBtn);
        }
    }
    else {
        template = document.createElement('div');
        template.className = 'experience-row';
        template.style.cssText = 'margin-bottom: 20px; padding: 16px; border: 1px solid var(--border); border-radius: 12px; background: var(--surface);';
        template.innerHTML = `
            <div class="form-grid">
                <label>Company<input name="experience_company[]" placeholder="ABC Technologies Ltd." /></label>
                <label>Role<input name="experience_role[]" placeholder="Software Engineer" /></label>
                <label>Duration<input name="experience_period[]" placeholder="Jan 2023 - Present" /></label>
                <label>Responsibilities<textarea name="experience_responsibilities[]" rows="4" placeholder="Developed and maintained REST APIs using Django REST Framework&#10;Implemented CI/CD pipelines using GitHub Actions&#10;Collaborated with cross-functional teams to deliver features on time"></textarea></label>
            </div>
            <button type="button" class="btn btn-danger btn-sm delete-row-btn" style="margin-top: 12px; padding: 8px 12px;" onclick="deleteRow(this)">Delete</button>
        `;
    }

    container.appendChild(template);
}

function addEducationRow() {
    const container = document.getElementById('education-rows');
    let template;
    
    if (container.children.length > 0) {
        template = container.children[0].cloneNode(true);
        template.querySelectorAll('input, textarea').forEach(i => { i.value = ''; });
        
        // add delete button
        if (!template.querySelector('.delete-row-btn')) {
            const deleteBtn = document.createElement('button');
            deleteBtn.type = 'button';
            deleteBtn.className = 'btn btn-danger btn-sm delete-row-btn';
            deleteBtn.style.cssText = 'margin-top: 12px; padding: 8px 12px;';
            deleteBtn.textContent = 'Delete';
            deleteBtn.onclick = function() { deleteRow(this); };
            template.appendChild(deleteBtn);
        }
    }
    else {
        template = document.createElement('div');
        template.className = 'edu-row';
        template.style.cssText = 'margin-bottom: 20px; padding: 16px; border: 1px solid var(--border); border-radius: 12px; background: var(--surface);';
        template.innerHTML = `
            <div class="form-grid">
                <label>Degree<input name="education_degree[]" placeholder="B.Sc in Computer Science and Engineering" /></label>
                <label>Institution<input name="education_institution[]" placeholder="National University, Bangladesh" /></label>
                <label>Period<input name="education_period[]" placeholder="2020 - 2024" /></label>
                <label>GPA/CGPA<input name="education_gpa_cgpa[]" placeholder="3.75" /></label>
                <label>Max GPA<input name="education_gpa_max[]" placeholder="4.00" /></label>
                <label>Details<textarea name="education_details[]" rows="2" placeholder="Relevant coursework: Data Structures, Algorithms, Machine Learning, Database Systems"></textarea></label>
            </div>
            <button type="button" class="btn btn-danger btn-sm delete-row-btn" style="margin-top: 12px; padding: 8px 12px;" onclick="deleteRow(this)">Delete</button>
        `;
    }

    container.appendChild(template);
}

function addCertificationRow() {
    const container = document.getElementById('certification-rows');
    let template;
    
    if (container.children.length > 0) {
        template = container.children[0].cloneNode(true);
        template.querySelectorAll('input, textarea').forEach(i => { i.value = ''; });
        
        // add delete button
        if (!template.querySelector('.delete-row-btn')) {
            const deleteBtn = document.createElement('button');
            deleteBtn.type = 'button';
            deleteBtn.className = 'btn btn-danger btn-sm delete-row-btn';
            deleteBtn.style.cssText = 'margin-top: 12px; padding: 8px 12px;';
            deleteBtn.textContent = 'Delete';
            deleteBtn.onclick = function() { deleteRow(this); };
            template.appendChild(deleteBtn);
        }
    }
    else {
        template = document.createElement('div');
        template.className = 'certification-row';
        template.style.cssText = 'margin-bottom: 20px; padding: 16px; border: 1px solid var(--border); border-radius: 12px; background: var(--surface);';
        template.innerHTML = `
            <div class="form-grid">
                <label>Certification Name<input name="certification_name[]" placeholder="AWS Certified Solutions Architect" /></label>
                <label>Issuing Organization<input name="certification_issuer[]" placeholder="Amazon Web Services" /></label>
                <label>Date Obtained<input name="certification_date[]" placeholder="Jan 2024" /></label>
                <label>Expiration Date<input name="certification_expiration[]" placeholder="Jan 2027 or Does not expire" /></label>
                <label>Credential ID<input name="certification_credential_id[]" placeholder="ABC123XYZ (optional)" /></label>
                <label>Credential URL<input name="certification_credential_url[]" placeholder="https://www.credly.com/badges/... (optional)" /></label>
            </div>
            <button type="button" class="btn btn-danger btn-sm delete-row-btn" style="margin-top: 12px; padding: 8px 12px;" onclick="deleteRow(this)">Delete</button>
        `;
    }

    container.appendChild(template);
}

function addHonorRow() {
    const container = document.getElementById('honor-rows');
    let template;
    
    if (container.children.length > 0) {
        template = container.children[0].cloneNode(true);
        template.querySelectorAll('input, textarea').forEach(i => { i.value = ''; });
        
        // add delete button if not present
        if (!template.querySelector('.delete-row-btn')) {
            const deleteBtn = document.createElement('button');
            deleteBtn.type = 'button';
            deleteBtn.className = 'btn btn-danger btn-sm delete-row-btn';
            deleteBtn.style.cssText = 'margin-top: 12px; padding: 8px 12px;';
            deleteBtn.textContent = 'Delete';
            deleteBtn.onclick = function() { deleteRow(this); };
            template.appendChild(deleteBtn);
        }
    }
    else {
        template = document.createElement('div');
        template.className = 'honor-row';
        template.style.cssText = 'margin-bottom: 20px; padding: 16px; border: 1px solid var(--border); border-radius: 12px; background: var(--surface);';
        template.innerHTML = `
            <div class="form-grid">
                <label>Title<input name="honor_title[]" placeholder="Dean's List Award" /></label>
                <label>Issuer<input name="honor_issuer[]" placeholder="National University, Bangladesh" /></label>
                <label>Date Received<input name="honor_date[]" placeholder="June 2023" /></label>
                <label>Description<textarea name="honor_description[]" rows="2" placeholder="Awarded for academic excellence with CGPA 3.9/4.0"></textarea></label>
            </div>
            <button type="button" class="btn btn-danger btn-sm delete-row-btn" style="margin-top: 12px; padding: 8px 12px;" onclick="deleteRow(this)">Delete</button>
        `;
    }

    container.appendChild(template);
}

// delete row function
function deleteRow(button) {
    const row = button.closest('.skill-category-row, .project-row, .experience-row, .edu-row, .certification-row, .honor-row');
    if (row) {
        row.remove();
    }
}

// combobox initialization and handling
function initCombobox(input) {
    const wrapper = input.closest('.combobox-wrapper');
    const dropdown = wrapper.querySelector('.combobox-dropdown');
    const options = wrapper.querySelectorAll('.combobox-option');
    const hiddenInput = wrapper.querySelector('.skill-custom-category');
    let highlightedIndex = -1;

    // toggle dropdown
    input.addEventListener('focus', () => {
        dropdown.classList.add('open');
        input.setAttribute('aria-expanded', 'true');
        filterOptions('');
    });

    input.addEventListener('click', (e) => {
        e.stopPropagation();
        dropdown.classList.add('open');
        input.setAttribute('aria-expanded', 'true');
        filterOptions(input.value);
    });

    // filter options on input
    input.addEventListener('input', () => {
        filterOptions(input.value);
    });

    // keyboard navigation
    input.addEventListener('keydown', (e) => {
        if (!dropdown.classList.contains('open')) return;
        
        const visibleOptions = Array.from(options).filter(opt => opt.style.display !== 'none');
        
        if (e.key === 'ArrowDown') {
            e.preventDefault();
            highlightedIndex = Math.min(highlightedIndex + 1, visibleOptions.length - 1);
            highlightOption(visibleOptions);
        }
        else if (e.key === 'ArrowUp') {
            e.preventDefault();
            highlightedIndex = Math.max(highlightedIndex - 1, -1);
            highlightOption(visibleOptions);
        }
        else if (e.key === 'Enter') {
            e.preventDefault();
            if (highlightedIndex >= 0 && visibleOptions[highlightedIndex]) {
                selectOption(visibleOptions[highlightedIndex]);
            }
            else if (input.value.trim()) {
                // Allow custom input
                setCustomValue(input.value.trim());
            }
        }
        else if (e.key === 'Escape') {
            closeDropdown();
        }
        else if (e.key === 'Tab') {
            closeDropdown();
        }
    });

    // click outside to close
    document.addEventListener('click', (e) => {
        if (!wrapper.contains(e.target)) {
            closeDropdown();
        }
    });

    // pption click handlers
    options.forEach(option => {
        option.addEventListener('click', () => selectOption(option));
        option.addEventListener('mouseenter', () => {
            options.forEach(o => o.classList.remove('highlighted'));
            option.classList.add('highlighted');
        });
    });

    function filterOptions(query) {
        const lowerQuery = query.toLowerCase();
        let hasVisibleOptions = false;
        
        options.forEach((option, index) => {
            const text = option.textContent.toLowerCase();
            const isCustom = option.classList.contains('combobox-custom');
            
            if (isCustom || text.includes(lowerQuery)) {
                option.style.display = '';
                hasVisibleOptions = true;
            }
            else {
                option.style.display = 'none';
            }
        });

        highlightedIndex = -1;
        dropdown.classList.toggle('open', hasVisibleOptions);
        input.setAttribute('aria-expanded', hasVisibleOptions);
    }

    function highlightOption(visibleOptions) {
        visibleOptions.forEach((opt, idx) => {
            opt.classList.toggle('highlighted', idx === highlightedIndex);
        });
    }

    function selectOption(option) {
        const value = option.dataset.value;
        const isCustom = option.classList.contains('combobox-custom');
        
        if (isCustom || value === 'custom') {
            // Show input for custom value
            input.value = '';
            input.placeholder = 'Type custom category...';
            hiddenInput.value = '';
            closeDropdown();
            input.focus();
        }
        else {
            input.value = value;
            hiddenInput.value = '';
            closeDropdown();
        }
    }

    function setCustomValue(value) {
        input.value = value;
        hiddenInput.value = value;
        closeDropdown();
    }

    function closeDropdown() {
        dropdown.classList.remove('open');
        input.setAttribute('aria-expanded', 'false');
        highlightedIndex = -1;
        options.forEach(o => o.classList.remove('highlighted'));
    }

    // initialize: if input has value that's not in default categories, set as custom
    const initialValue = input.value.trim();
    if (initialValue && !Array.from(options).some(o => o.dataset.value === initialValue && !o.classList.contains('combobox-custom'))) {
        hiddenInput.value = initialValue;
    }
}

// initialize comboboxes on DOMContentLoaded
document.addEventListener('DOMContentLoaded', function() {
    document.querySelectorAll('.skill-category-combobox').forEach(input => {
        initCombobox(input);
    });
    document.querySelectorAll('.project-category-combobox').forEach(input => {
        initCombobox(input);
    });
});