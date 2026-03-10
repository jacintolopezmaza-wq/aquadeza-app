// App State Manager
const App = {
    currentUser: null,
    currentStation: null,
    currentBotosTab: 'hoja1',
    currentBombeoBotosTab: 'hoja1',
    currentCatasosTab: 'hoja1',
    currentEtapTab: 'hoja1',
    currentCorredoiraTab: 'diario',
    currentVilatuxeTab: 'hoja1',
    generateYearOptions(selectedYear) {
        let html = '';
        for (let y = 2024; y <= 2050; y++) {
            html += `<option value="${y}" ${y === selectedYear ? 'selected' : ''}>${y}</option>`;
        }
        return html;
    },
    views: {
        login: document.getElementById('login-view'),
        dashboard: document.getElementById('dashboard-view'),
        form: document.getElementById('form-view')
    },

    DataManager: {
        saveEntry(station, data) {
            try {
                const date = data.fecha || new Date().toISOString().split('T')[0];
                const logs = JSON.parse(localStorage.getItem(`logs_${station}`) || '[]');
                const index = logs.findIndex(l => l.fecha === date);
                const entry = { ...data, fecha: date };

                if (index > -1) {
                    // Merge existing data with new data to prevent data loss across tabs
                    logs[index] = { ...logs[index], ...entry };
                } else {
                    logs.push(entry);
                }

                // Robust sort handling potential missing dates
                logs.sort((a, b) => (a.fecha || '').localeCompare(b.fecha || ''));
                localStorage.setItem(`logs_${station}`, JSON.stringify(logs));
            } catch (e) {
                console.error("Error saving data:", e);
                alert("Hubo un error al guardar los datos.");
            }
        },

        getPreviousEntry(station, currentDate) {
            const logs = JSON.parse(localStorage.getItem(`logs_${station}`) || '[]');
            return logs.filter(l => l.fecha < currentDate).pop() || null;
        }
    },

    init() {
        this.setupEventListeners();
        this.checkAuth();
        this.registerServiceWorker();
    },

    setupEventListeners() {
        document.getElementById('logout-btn').addEventListener('click', () => this.handleLogout());

        document.getElementById('login-manual-btn').addEventListener('click', () => {
            const nameInput = document.getElementById('login-name');
            if (nameInput.value.trim()) {
                this.handleLogin(nameInput.value.trim());
            } else {
                alert("Por favor, ingrese su nombre.");
            }
        });

        document.querySelectorAll('.station-card').forEach(card => {
            card.addEventListener('click', () => this.showStationForm(card.getAttribute('data-station')));
        });

        document.getElementById('back-btn').addEventListener('click', () => this.showView('dashboard'));

        document.getElementById('station-form').addEventListener('submit', (e) => {
            e.preventDefault();
            this.handleFormSubmit();
        });

        // Export Listeners
        document.getElementById('export-excel-btn').addEventListener('click', () => this.exportAllData('excel'));
        document.getElementById('export-pdf-btn').addEventListener('click', () => this.exportAllData('pdf'));
    },


    checkAuth() {
        const savedUser = localStorage.getItem('control_user');
        if (savedUser) {
            this.currentUser = JSON.parse(savedUser);
            this.showView('dashboard');
        } else {
            this.showView('login');
        }
    },

    handleLogin(username) {
        this.currentUser = { name: username };
        localStorage.setItem('control_user', JSON.stringify(this.currentUser));
        this.showView('dashboard');
    },

    handleLogout() {
        localStorage.removeItem('control_user');
        this.currentUser = null;
        this.showView('login');
    },

    showView(viewName) {
        Object.keys(this.views).forEach(v => {
            this.views[v].style.display = v === viewName ? (v === 'login' ? 'flex' : 'block') : 'none';
        });

        if (viewName === 'dashboard' && this.currentUser) {
            document.getElementById('welcome-user').textContent = `Hola, ${this.currentUser.name}`;
            this.checkMonthlyBackup();
        }
    },

    checkMonthlyBackup() {
        const lastBackup = localStorage.getItem('last_backup_month');
        const currentMonth = new Date().getMonth();
        if (lastBackup !== null && parseInt(lastBackup) !== currentMonth) {
            alert("¡Es un nuevo mes! Recuerda realizar la copia de seguridad del mes anterior pulsando los botones de Exportar.");
            localStorage.setItem('last_backup_month', currentMonth);
        } else if (lastBackup === null) {
            localStorage.setItem('last_backup_month', currentMonth);
        }
    },

    registerServiceWorker() {
        if ('serviceWorker' in navigator) {
            navigator.serviceWorker.register('sw.js')
                .then(reg => console.log('Service Worker registrado'))
                .catch(err => console.log('Error registrando SW', err));
        }
    },

    showStationForm(station) {
        this.currentStation = station;
        const titleMap = {
            'ETAP': 'ETAP',
            'CATASOS': 'CATASOS',
            'BOMBEO_BOTOS': 'BOMBEO DE BOTOS',
            'EDAR_BOTOS': 'EDAR DE BOTOS',
            'EDAR_CORREDOIRA': 'EDAR CORREDOIRA',
            'VILATUXE': 'VILATUXE'
        };
        document.getElementById('station-title').textContent = titleMap[station] || station;
        this.renderFormFields(station);
        this.showView('form');
    },

    renderFormFields(station) {
        const fieldContainer = document.getElementById('form-fields');
        const formElement = document.getElementById('station-form');
        fieldContainer.innerHTML = '';

        if (station === 'EDAR_BOTOS' || station === 'EDAR_CORREDOIRA' || station === 'ETAP' || station === 'BOMBEO_BOTOS' || station === 'CATASOS' || station === 'VILATUXE') {
            formElement.classList.add('sheet-form');
            if (station === 'EDAR_BOTOS') {
                this.renderEdarBotosSheet(fieldContainer);
            } else if (station === 'ETAP') {
                this.renderEtapSheet(fieldContainer);
            } else if (station === 'BOMBEO_BOTOS') {
                this.renderBombeoBotosSheet(fieldContainer);
            } else if (station === 'CATASOS') {
                this.renderCatasosSheet(fieldContainer);
            } else if (station === 'VILATUXE') {
                this.renderVilatuxeSheet(fieldContainer);
            } else {
                this.renderEdarCorredoiraSheet(fieldContainer);
            }
        } else {
            formElement.classList.remove('sheet-form');
            fieldContainer.innerHTML = `
                <div class="input-group">
                    <label>📅 Fecha (automática)</label>
                    <input type="date" id="form-date" value="${new Date().toISOString().split('T')[0]}" readonly style="pointer-events:none; opacity:0.75; cursor:not-allowed; background:#1a1a2e; border: 1px solid #444;">
                </div>
                <div class="input-group">
                    <label>Observaciones Generales</label>
                    <textarea id="form-obs" style="width:100%; height:100px; padding:12px; background:#2c2c2e; border:1px solid var(--border); border-radius:var(--radius); color:white;"></textarea>
                </div>
            `;
        }
    },

    renderSection(title, fields, previousEntry, isDouble = false) {
        let html = `<h2 class="section-title">${title}</h2>`;
        html += `
            <div class="table-header" style="grid-template-columns: ${isDouble ? '2fr 1fr 1fr 1fr' : '2fr 1fr 1fr'};">
                <span>Parámetro</span>
                ${isDouble ? '<span>Lectura</span>' : ''}
                <span>${isDouble ? 'Horas' : 'Valor'}</span>
                <span>Resta</span>
            </div>
            <div class="double-input-grid">
        `;

        fields.forEach(f => {
            const name = f.name || f.label.toLowerCase().replace(/ /g, '_').replace(/[()°/]/g, '');
            const prevVal = previousEntry ? (previousEntry[name] || 0) : 0;

            if (isDouble) {
                html += `
                    <div class="double-field-row" data-field="${name}" data-type="double">
                        <label>${f.label}</label>
                        <input type="number" class="lect-input" placeholder="Lect." data-prev="${prevVal.lect || 0}">
                        <input type="number" class="hrs-input" placeholder="Hrs." data-prev="${prevVal.hrs || 0}">
                        <div class="diff-display">0</div>
                    </div>
                `;
            } else {
                html += `
                    <div class="double-field-row" data-field="${name}" data-type="single" style="grid-template-columns: 2fr 1fr 1fr;">
                        <label>${f.label}</label>
                        <input type="number" class="val-input" placeholder="Valor" data-prev="${prevVal}">
                        <div class="diff-display">0</div>
                    </div>
                `;
            }
        });

        html += `</div>`;
        return html;
    },

    renderEdarBotosSheet(container) {
        if (!this.currentMonth) {
            const now = new Date();
            this.currentMonth = now.getMonth();
            this.currentYear = now.getFullYear();
        }

        const months = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"];

        // Metadata Inline Controls
        let html = `
            <div class="premium-header">
                <div class="header-logo-top">AQUADEZA</div>
                <div class="header-main-titles">
                    <div class="brand-tagline">SERVICIO DE LALÍN</div>
                    <h1 class="station-title">EDAR DE BOTOS</h1>
                    <div class="sheet-info-badge">${this.currentBotosTab.toUpperCase()}</div>
                </div>
                <div class="header-controls-row">
                    <div class="control-pill">
                        <label>MES</label>
                        <select id="month-select">
                            ${months.map((m, i) => `<option value="${i}" ${i === this.currentMonth ? 'selected' : ''}>${m}</option>`).join('')}
                        </select>
                    </div>
                    <div class="control-pill">
                        <label>AÑO</label>
                        <select id="year-select">
                            ${this.generateYearOptions(this.currentYear)}
                        </select>
                    </div>
                </div>
            </div>
            
            <div class="sheet-tabs">
                <button class="tab-btn ${this.currentBotosTab === 'hoja1' ? 'active' : ''}" data-tab="hoja1">HOJA 1 (DIARIO)</button>
                <button class="tab-btn ${this.currentBotosTab === 'hoja2' ? 'active' : ''}" data-tab="hoja2">HOJA 2 (HORAS 1)</button>
                <button class="tab-btn ${this.currentBotosTab === 'hoja3' ? 'active' : ''}" data-tab="hoja3">HOJA 3 (HORAS 2)</button>
                <button class="tab-btn ${this.currentBotosTab === 'hoja4' ? 'active' : ''}" data-tab="hoja4">HOJA 4 (HORAS 3)</button>
                <button class="tab-btn ${this.currentBotosTab === 'hoja_energia' ? 'active' : ''}" data-tab="hoja_energia">CUADRO ENERGÍA</button>
            </div>
            <div id="botos-sheet-content">
                ${this.currentBotosTab === 'hoja1' ? this.renderEdarBotosHoja1Content() : ''}
                ${this.currentBotosTab === 'hoja2' ? this.renderEdarBotosHoja2Content() : ''}
                ${this.currentBotosTab === 'hoja3' ? this.renderEdarBotosHoja3Content() : ''}
                ${this.currentBotosTab === 'hoja4' ? this.renderEdarBotosHoja4Content() : ''}
                ${this.currentBotosTab === 'hoja_energia' ? this.renderEdarBotosEnergiaContent() : ''}
            </div>
            <div style="margin-top:20px; text-align:right; padding-right:10px;">
                <button type="submit" class="btn-primary" style="width:auto; padding:10px 30px;">GUARDAR TODO EL MES</button>
            </div>
            <div class="aguadeza-footer">AQUADEZA</div>
        `;
        container.innerHTML = html;

        this.setupMonthlyEventListeners();
        this.setupBotosTabListeners();
        this.recalculateDailyConsumption();
        this.recalculateTotals();
    },



    setupBotosTabListeners() {
        document.querySelectorAll('.tab-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                this.currentBotosTab = e.target.getAttribute('data-tab');
                this.renderEdarBotosSheet(document.getElementById('form-fields'));
            });
        });
    },

    renderBombeoBotosSheet(container) {
        if (!this.currentMonth) {
            const now = new Date();
            this.currentMonth = now.getMonth();
            this.currentYear = now.getFullYear();
        }
        const months = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"];

        let html = `
            <div class="premium-header">
                <div class="header-row">
                    <div class="header-brand">
                        <span class="brand-name">AQUADEZA</span>
                        <span class="brand-tagline">SERVICIO DE LALÍN</span>
                    </div>
                    <div class="header-station-box">
                        <h1 class="station-title">BOMBEO DE BOTOS</h1>
                        <div class="sheet-info-badge">HOJA: ${this.currentBombeoBotosTab === 'hoja1' ? '1' : '2'}</div>
                    </div>
                </div>
                <div class="header-controls-row">
                    <div class="control-pill">
                        <label>MES</label>
                        <select id="month-select">
                            ${months.map((m, i) => `<option value="${i}" ${i === this.currentMonth ? 'selected' : ''}>${m}</option>`).join('')}
                        </select>
                    </div>
                    <div class="control-pill">
                        <label>AÑO</label>
                        <select id="year-select">
                            ${this.generateYearOptions(this.currentYear)}
                        </select>
                    </div>
                </div>
            </div>

            <div class="sheet-tabs">
                <button class="tab-btn ${this.currentBombeoBotosTab === 'hoja1' ? 'active' : ''}" data-tab="hoja1">HOJA 1 (BOMBAS/ACTIVA)</button>
                <button class="tab-btn ${this.currentBombeoBotosTab === 'hoja2' ? 'active' : ''}" data-tab="hoja2">HOJA 2 (REACTIVA/MAXIMETRO)</button>
            </div>

            <div id="botos-sheet-content">
                ${this.currentBombeoBotosTab === 'hoja1' ? this.renderBombeoBotosHoja1Content() : ''}
                ${this.currentBombeoBotosTab === 'hoja2' ? this.renderBombeoBotosHoja2Content() : ''}
            </div>

            <div style="margin-top:20px; text-align:right; padding-right:10px;">
                <button type="submit" class="btn-primary" style="width:auto; padding:10px 30px;">GUARDAR TODO EL MES</button>
            </div>
        `;
        container.innerHTML = html;
        this.setupMonthlyEventListeners();
        this.setupBombeoBotosTabListeners();
        this.recalculateTotals();
    },

    setupBombeoBotosTabListeners() {
        document.querySelectorAll('.tab-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                this.currentBombeoBotosTab = e.target.getAttribute('data-tab');
                this.renderBombeoBotosSheet(document.getElementById('form-fields'));
            });
        });
    },

    renderBombeoBotosHoja1Content() {
        const logs = JSON.parse(localStorage.getItem(`logs_BOMBEO_BOTOS`) || '[]');
        const monthKey = `${this.currentYear}-${String(this.currentMonth + 1).padStart(2, '0')}`;

        let html = `<div class="sheet-table bombeo-botos-hoja1">
            <div class="sheet-cell sheet-header-cell span-row-2">DIA</div>
            <div class="sheet-cell sheet-header-cell span-row-2" style="display:none;">HORA</div>
            <div class="sheet-cell sheet-header-cell span-col-6">HORAS BOMBEO</div>
            <div class="sheet-cell sheet-header-cell span-col-7">ENERGIA ACTIVA</div>

            <div class="sheet-cell sheet-header-cell">BOMBA 1</div>
            <div class="sheet-cell sheet-header-cell">DIF</div>
            <div class="sheet-cell sheet-header-cell">BOMBA 2</div>
            <div class="sheet-cell sheet-header-cell">DIF</div>
            <div class="sheet-cell sheet-header-cell">BOMBA 3</div>
            <div class="sheet-cell sheet-header-cell">DIF</div>
            <div class="sheet-cell sheet-header-cell">P 1.18.1</div>
            <div class="sheet-cell sheet-header-cell">P 1.18.2</div>
            <div class="sheet-cell sheet-header-cell">P 1.18.3</div>
            <div class="sheet-cell sheet-header-cell">P 1.18.4</div>
            <div class="sheet-cell sheet-header-cell">P 1.18.5</div>
            <div class="sheet-cell sheet-header-cell">P 1.18.6</div>
            <div class="sheet-cell sheet-header-cell">P 1.18.0</div>
        `;

        const daysInMonth = new Date(this.currentYear, this.currentMonth + 1, 0).getDate();
        for (let d = 0; d <= 31; d++) {
            const dateStr = `${monthKey}-${String(d).padStart(2, '0')}`;
            const log = logs.find(l => l.fecha === dateStr) || {};
            const isInvalidDay = d > daysInMonth;
            const isInitial = d === 0;

            html += `
                <div class="sheet-cell ${isInitial ? 'initial-row-cell' : ''} ${isInvalidDay ? 'disabled-day' : ''}">${d}</div>
                <div class="sheet-cell" style="display:none;"><input type="time" class="row-input" data-date="${dateStr}" data-field="hora" value="${log.hora || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="b1" value="${log.b1 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="b1_dif" value="${log.b1_dif || ''}" disabled></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="b2" value="${log.b2 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="b2_dif" value="${log.b2_dif || ''}" disabled></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="b3" value="${log.b3 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="b3_dif" value="${log.b3_dif || ''}" disabled></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="p1181" value="${log.p1181 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="p1182" value="${log.p1182 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="p1183" value="${log.p1183 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="p1184" value="${log.p1184 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="p1185" value="${log.p1185 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="p1186" value="${log.p1186 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="p1180" value="${log.p1180 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
            `;
        }
        return html + `</div>`;
    },
    renderBombeoBotosHoja2Content() {
        const logs = JSON.parse(localStorage.getItem(`logs_BOMBEO_BOTOS`) || '[]');
        const monthKey = `${this.currentYear}-${String(this.currentMonth + 1).padStart(2, '0')}`;

        let html = `<div class="sheet-table bombeo-botos-hoja2">
            <div class="sheet-cell sheet-header-cell span-row-2">DIA</div>
            <div class="sheet-cell sheet-header-cell span-row-2" style="display:none;">HORA</div>
            <div class="sheet-cell sheet-header-cell span-col-7">ENERGIA REACTIVA</div>
            <div class="sheet-cell sheet-header-cell span-col-7">MAXIMETRO</div>

            <div class="sheet-cell sheet-header-cell">P 1.58.1</div>
            <div class="sheet-cell sheet-header-cell">P 1.58.2</div>
            <div class="sheet-cell sheet-header-cell">P 1.58.3</div>
            <div class="sheet-cell sheet-header-cell">P 1.58.4</div>
            <div class="sheet-cell sheet-header-cell">P 1.58.5</div>
            <div class="sheet-cell sheet-header-cell">P 1.58.6</div>
            <div class="sheet-cell sheet-header-cell">P 1.58.0</div>
            <div class="sheet-cell sheet-header-cell">P 1.16.1</div>
            <div class="sheet-cell sheet-header-cell">P 1.16.2</div>
            <div class="sheet-cell sheet-header-cell">P 1.16.3</div>
            <div class="sheet-cell sheet-header-cell">P 1.16.4</div>
            <div class="sheet-cell sheet-header-cell">P 1.16.5</div>
            <div class="sheet-cell sheet-header-cell">P 1.16.6</div>
            <div class="sheet-cell sheet-header-cell">P 1.16.0</div>
        `;

        const daysInMonth = new Date(this.currentYear, this.currentMonth + 1, 0).getDate();
        for (let d = 0; d <= 31; d++) {
            const dateStr = `${monthKey}-${String(d).padStart(2, '0')}`;
            const log = logs.find(l => l.fecha === dateStr) || {};
            const isInvalidDay = d > daysInMonth;
            const isInitial = d === 0;

            html += `
                <div class="sheet-cell ${isInitial ? 'initial-row-cell' : ''} ${isInvalidDay ? 'disabled-day' : ''}">${d}</div>
                <div class="sheet-cell" style="display:none;"><input type="time" class="row-input" data-date="${dateStr}" data-field="hora" value="${log.hora || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="p1581" value="${log.p1581 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="p1582" value="${log.p1582 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="p1583" value="${log.p1583 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="p1584" value="${log.p1584 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="p1585" value="${log.p1585 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="p1586" value="${log.p1586 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="p1580" value="${log.p1580 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="p1161" value="${log.p1161 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="p1162" value="${log.p1162 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="p1163" value="${log.p1163 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="p1164" value="${log.p1164 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="p1165" value="${log.p1165 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="p1166" value="${log.p1166 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="p1160" value="${log.p1160 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
            `;
        }
        return html + `</div>`;
    },


    renderCatasosSheet(container) {
        if (!this.currentMonth) {
            const now = new Date();
            this.currentMonth = now.getMonth();
            this.currentYear = now.getFullYear();
        }
        const months = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"];

        let html = `
            <div class="premium-header">
                <div class="header-row">
                    <div class="header-brand">
                        <span class="brand-name">AQUADEZA</span>
                        <span class="brand-tagline">SERVICIO DE LALÍN</span>
                    </div>
                    <div class="header-station-box">
                        <h1 class="station-title">BOMBEO DE CATASOS</h1>
                        <div class="sheet-info-badge">HOJA: ${this.currentCatasosTab === 'hoja1' ? '1' : '2'}</div>
                    </div>
                </div>
                <div class="header-controls-row">
                    <div class="control-pill">
                        <label>MES</label>
                        <select id="month-select">
                            ${months.map((m, i) => `<option value="${i}" ${i === this.currentMonth ? 'selected' : ''}>${m}</option>`).join('')}
                        </select>
                    </div>
                    <div class="control-pill">
                        <label>AÑO</label>
                        <select id="year-select">
                            ${this.generateYearOptions(this.currentYear)}
                        </select>
                    </div>
                </div>
            </div>

            <div class="sheet-tabs">
                <button class="tab-btn ${this.currentCatasosTab === 'hoja1' ? 'active' : ''}" data-tab="hoja1">HOJA 1 (BOMBAS/ACTIVA)</button>
                <button class="tab-btn ${this.currentCatasosTab === 'hoja2' ? 'active' : ''}" data-tab="hoja2">HOJA 2 (REACTIVA/MAXIMETRO)</button>
            </div>

            <div id="catasos-sheet-content">
                ${this.currentCatasosTab === 'hoja1' ? this.renderCatasosHoja1Content() : ''}
                ${this.currentCatasosTab === 'hoja2' ? this.renderCatasosHoja2Content() : ''}
            </div>

            <div style="margin-top:20px; text-align:right; padding-right:10px;">
                <button type="submit" class="btn-primary" style="width:auto; padding:10px 30px;">GUARDAR TODO EL MES</button>
            </div>
        `;
        container.innerHTML = html;
        this.setupMonthlyEventListeners();
        this.setupCatasosTabListeners();
        this.recalculateTotals();
    },

    setupCatasosTabListeners() {
        document.querySelectorAll('.tab-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                this.currentCatasosTab = e.target.getAttribute('data-tab');
                this.renderCatasosSheet(document.getElementById('form-fields'));
            });
        });
    },

    renderCatasosHoja1Content() {
        const logs = JSON.parse(localStorage.getItem(`logs_CATASOS`) || '[]');
        const monthKey = `${this.currentYear}-${String(this.currentMonth + 1).padStart(2, '0')}`;

        let html = `<div class="sheet-table bombeo-catasos-hoja1">
            <div class="sheet-cell sheet-header-cell span-row-2">DIA</div>
            <div class="sheet-cell sheet-header-cell span-row-2" style="display:none;">HORA</div>
            <div class="sheet-cell sheet-header-cell span-col-6">HORAS BOMBEO</div>
            <div class="sheet-cell sheet-header-cell span-col-7">ENERGIA ACTIVA</div>

            <div class="sheet-cell sheet-header-cell">BOMBA 1</div>
            <div class="sheet-cell sheet-header-cell">DIF</div>
            <div class="sheet-cell sheet-header-cell">BOMBA 2</div>
            <div class="sheet-cell sheet-header-cell">DIF</div>
            <div class="sheet-cell sheet-header-cell">BOMBA 3</div>
            <div class="sheet-cell sheet-header-cell">DIF</div>
            <div class="sheet-cell sheet-header-cell">P 1.18.1</div>
            <div class="sheet-cell sheet-header-cell">P 1.18.2</div>
            <div class="sheet-cell sheet-header-cell">P 1.18.3</div>
            <div class="sheet-cell sheet-header-cell">P 1.18.4</div>
            <div class="sheet-cell sheet-header-cell">P 1.18.5</div>
            <div class="sheet-cell sheet-header-cell">P 1.18.6</div>
            <div class="sheet-cell sheet-header-cell">P 1.18.0</div>
        `;

        const daysInMonth = new Date(this.currentYear, this.currentMonth + 1, 0).getDate();
        for (let d = 0; d <= 31; d++) {
            const dateStr = `${monthKey}-${String(d).padStart(2, '0')}`;
            const log = logs.find(l => l.fecha === dateStr) || {};
            const isInvalidDay = d > daysInMonth;
            const isInitial = d === 0;

            html += `
                <div class="sheet-cell ${isInitial ? 'initial-row-cell' : ''} ${isInvalidDay ? 'disabled-day' : ''}">${d}</div>
                <div class="sheet-cell" style="display:none;"><input type="time" class="row-input" data-date="${dateStr}" data-field="hora" value="${log.hora || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="b1" value="${log.b1 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="b1_dif" value="${log.b1_dif || ''}" disabled></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="b2" value="${log.b2 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="b2_dif" value="${log.b2_dif || ''}" disabled></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="b3" value="${log.b3 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="b3_dif" value="${log.b3_dif || ''}" disabled></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="p1181" value="${log.p1181 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="p1182" value="${log.p1182 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="p1183" value="${log.p1183 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="p1184" value="${log.p1184 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="p1185" value="${log.p1185 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="p1186" value="${log.p1186 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="p1180" value="${log.p1180 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
            `;
        }
        return html + `</div>`;
    },

    renderCatasosHoja2Content() {
        const logs = JSON.parse(localStorage.getItem(`logs_CATASOS`) || '[]');
        const monthKey = `${this.currentYear}-${String(this.currentMonth + 1).padStart(2, '0')}`;

        let html = `<div class="sheet-table bombeo-catasos-hoja2">
            <div class="sheet-cell sheet-header-cell span-row-2">DIA</div>
            <div class="sheet-cell sheet-header-cell span-row-2" style="display:none;">HORA</div>
            <div class="sheet-cell sheet-header-cell span-col-7">ENERGIA REACTIVA</div>
            <div class="sheet-cell sheet-header-cell span-col-7">MAXIMETRO</div>

            <div class="sheet-cell sheet-header-cell">P 1.58.1</div>
            <div class="sheet-cell sheet-header-cell">P 1.58.2</div>
            <div class="sheet-cell sheet-header-cell">P 1.58.3</div>
            <div class="sheet-cell sheet-header-cell">P 1.58.4</div>
            <div class="sheet-cell sheet-header-cell">P 1.58.5</div>
            <div class="sheet-cell sheet-header-cell">P 1.58.6</div>
            <div class="sheet-cell sheet-header-cell">P 1.58.0</div>
            <div class="sheet-cell sheet-header-cell">P 1.16.1</div>
            <div class="sheet-cell sheet-header-cell">P 1.16.2</div>
            <div class="sheet-cell sheet-header-cell">P 1.16.3</div>
            <div class="sheet-cell sheet-header-cell">P 1.16.4</div>
            <div class="sheet-cell sheet-header-cell">P 1.16.5</div>
            <div class="sheet-cell sheet-header-cell">P 1.16.6</div>
            <div class="sheet-cell sheet-header-cell">P 1.16.0</div>
        `;

        const daysInMonth = new Date(this.currentYear, this.currentMonth + 1, 0).getDate();
        for (let d = 0; d <= 31; d++) {
            const dateStr = `${monthKey}-${String(d).padStart(2, '0')}`;
            const log = logs.find(l => l.fecha === dateStr) || {};
            const isInvalidDay = d > daysInMonth;
            const isInitial = d === 0;

            html += `
                <div class="sheet-cell ${isInitial ? 'initial-row-cell' : ''} ${isInvalidDay ? 'disabled-day' : ''}">${d}</div>
                <div class="sheet-cell" style="display:none;"><input type="time" class="row-input" data-date="${dateStr}" data-field="hora" value="${log.hora || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="p1581" value="${log.p1581 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="p1582" value="${log.p1582 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="p1583" value="${log.p1583 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="p1584" value="${log.p1584 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="p1585" value="${log.p1585 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="p1586" value="${log.p1586 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="p1580" value="${log.p1580 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="p1161" value="${log.p1161 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="p1162" value="${log.p1162 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="p1163" value="${log.p1163 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="p1164" value="${log.p1164 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="p1165" value="${log.p1165 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="p1166" value="${log.p1166 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="p1160" value="${log.p1160 || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
            `;
        }
        return html + `</div>`;
    },

    saveSingleField(date, field, val) {

        const logs = JSON.parse(localStorage.getItem(`logs_${this.currentStation}`) || '[]');
        let log = logs.find(l => l.fecha === date);
        if (!log) {
            log = { fecha: date };
            logs.push(log);
        }
        const parsed = parseFloat(val);
        log[field] = (val === '' ? '' : (!isNaN(parsed) && isFinite(val) ? parsed : val));
        localStorage.setItem(`logs_${this.currentStation}`, JSON.stringify(logs));
    },

    renderEdarBotosHoja1Content() {
        let html = `<div class="sheet-table botos-diario">
                <!-- Grouped Headers Row -->
                <div class="sheet-cell sheet-header-cell span-row-2"><div class="header-horizontal">Día</div></div>
                <div class="sheet-cell sheet-header-cell span-row-2"><div class="header-horizontal">Hora</div></div>
                <div class="sheet-cell sheet-header-cell span-col-2">Tª</div>
                <div class="sheet-cell sheet-header-cell span-row-2"><div class="header-vertical">Precipitación mm.</div></div>
                <div class="sheet-cell sheet-header-cell span-col-2">CAUDAL</div>
                <div class="sheet-cell sheet-header-cell span-col-4">PARÁMETROS DE FUNCIONAMIENTO</div>
                <div class="sheet-cell sheet-header-cell span-col-4">CONSUMO ELÉCTRICO</div>
                <div class="sheet-cell sheet-header-cell span-row-2"><div class="header-horizontal">OBSERVACIONES</div></div>

                <!-- Sub-Headers Row -->
                <div class="sheet-cell sheet-header-cell"><div class="header-vertical">Aire ºC</div></div>
                <div class="sheet-cell sheet-header-cell"><div class="header-vertical">Agua ºC</div></div>
                <div class="sheet-cell sheet-header-cell"><div class="header-horizontal">Lectura</div></div>
                <div class="sheet-cell sheet-header-cell"><div class="header-horizontal">m³/día</div></div>
                <div class="sheet-cell sheet-header-cell"><div class="header-vertical">Desbaste m³</div></div>
                <div class="sheet-cell sheet-header-cell"><div class="header-vertical">Fangos filtro m³</div></div>
                <div class="sheet-cell sheet-header-cell"><div class="header-vertical">P A X cm.</div></div>
                <div class="sheet-cell sheet-header-cell"><div class="header-vertical">Polielectrolito kg.</div></div>
                <div class="sheet-cell sheet-header-cell"><div class="header-horizontal">PUNTA I</div></div>
                <div class="sheet-cell sheet-header-cell"><div class="header-horizontal">VALLE II</div></div>
                <div class="sheet-cell sheet-header-cell"><div class="header-horizontal">LLANO III</div></div>
                <div class="sheet-cell sheet-header-cell"><div class="header-horizontal">ENERGÍA REACTIVA</div></div>
        `;

        const logs = JSON.parse(localStorage.getItem(`logs_EDAR_BOTOS`) || '[]');
        const monthKey = `${this.currentYear}-${String(this.currentMonth + 1).padStart(2, '0')}`;
        const initialLog = logs.find(l => l.fecha === `${monthKey}-00`) || {};

        html += `
            <div class="sheet-cell initial-row-cell">-</div>
            <div class="sheet-cell" style="display:none;"><input type="time" class="row-input" data-date="${monthKey}-00" data-field="hora" value="${initialLog.hora || ''}" ></div>
            <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${monthKey}-00" data-field="taire" value="${initialLog.taire || ''}" ></div>
            <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${monthKey}-00" data-field="tagua" value="${initialLog.tagua || ''}" ></div>
            <div class="sheet-cell"><input type="text" inputmode="text" class="row-input" data-date="${monthKey}-00" data-field="precipitacion" value="${initialLog.precipitacion || ''}" ></div>
            <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${monthKey}-00" data-field="caudal_lect" value="${initialLog.caudal_lect || ''}" ></div>
            <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${monthKey}-00" data-field="caudal_m3d" value="${initialLog.caudal_m3d || ''}" ></div>
            <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${monthKey}-00" data-field="desbaste" value="${initialLog.desbaste || ''}" ></div>
            <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${monthKey}-00" data-field="fangos_filtro" value="${initialLog.fangos_filtro || ''}" ></div>
            <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${monthKey}-00" data-field="pax" value="${initialLog.pax || ''}" ></div>
            <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${monthKey}-00" data-field="polielectrolito" value="${initialLog.polielectrolito || ''}" ></div>
            <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${monthKey}-00" data-field="punta" value="${initialLog.punta || ''}" ></div>
            <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${monthKey}-00" data-field="valle" value="${initialLog.valle || ''}" ></div>
            <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${monthKey}-00" data-field="llano" value="${initialLog.llano || ''}" ></div>
            <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${monthKey}-00" data-field="reactiva" value="${initialLog.reactiva || ''}" ></div>
            <div class="sheet-cell"><input type="text" inputmode="text" class="row-input" data-date="${monthKey}-00" data-field="observaciones" value="${initialLog.observaciones || ''}" ></div>
        `;

        const daysInMonth = new Date(this.currentYear, this.currentMonth + 1, 0).getDate();
        for (let d = 1; d <= 31; d++) {
            const dateStr = `${this.currentYear}-${String(this.currentMonth + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
            const log = logs.find(l => l.fecha === dateStr) || {};
            const isInvalidDay = d > daysInMonth;

            html += `
                <div class="sheet-cell date-cell ${isInvalidDay ? 'disabled-row' : ''}">${d}</div>
                <div class="sheet-cell" style="display:none;"><input type="time" class="row-input" data-date="${dateStr}" data-field="hora" value="${log.hora || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${dateStr}" data-field="taire" value="${log.taire || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${dateStr}" data-field="tagua" value="${log.tagua || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="text" inputmode="text" class="row-input" data-date="${dateStr}" data-field="precipitacion" value="${log.precipitacion || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${dateStr}" data-field="caudal_lect" value="${log.caudal_lect || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${dateStr}" data-field="caudal_m3d" value="${log.caudal_m3d || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${dateStr}" data-field="desbaste" value="${log.desbaste || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${dateStr}" data-field="fangos_filtro" value="${log.fangos_filtro || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${dateStr}" data-field="pax" value="${log.pax || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${dateStr}" data-field="polielectrolito" value="${log.polielectrolito || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${dateStr}" data-field="punta" value="${log.punta || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${dateStr}" data-field="valle" value="${log.valle || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${dateStr}" data-field="llano" value="${log.llano || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${dateStr}" data-field="reactiva" value="${log.reactiva || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="text" inputmode="text" class="row-input" data-date="${dateStr}" data-field="observaciones" value="${log.observaciones || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
        `;
        }

        html += `
            <div class="sheet-cell footer-cell span-col-5" style="text-align:left; padding-left:10px;">TOTAL</div>
            <div class="sheet-cell footer-cell" id="total-caudal-lect">-</div><div class="sheet-cell footer-cell" id="total-caudal-m3d">-</div>
            <div class="sheet-cell footer-cell" id="total-desbaste">-</div><div class="sheet-cell footer-cell" id="total-fangos">-</div><div class="sheet-cell footer-cell" id="total-pax">-</div><div class="sheet-cell footer-cell" id="total-poli">-</div>
            <div class="sheet-cell footer-cell" id="total-punta">-</div><div class="sheet-cell footer-cell" id="total-valle">-</div><div class="sheet-cell footer-cell" id="total-llano">-</div><div class="sheet-cell footer-cell" id="total-reactiva">-</div>
            <div class="sheet-cell footer-cell"></div>

            <div class="sheet-cell footer-cell span-col-5" style="text-align:left; padding-left:10px;">MEDIA</div>
            <div class="sheet-cell footer-cell" id="media-caudal-lect">-</div><div class="sheet-cell footer-cell" id="media-caudal-m3d">-</div>
            <div class="sheet-cell footer-cell" id="media-desbaste">-</div><div class="sheet-cell footer-cell" id="media-fangos">-</div><div class="sheet-cell footer-cell" id="media-pax">-</div><div class="sheet-cell footer-cell" id="media-poli">-</div>
            <div class="sheet-cell footer-cell" id="media-punta">-</div><div class="sheet-cell footer-cell" id="media-valle">-</div><div class="sheet-cell footer-cell" id="media-llano">-</div><div class="sheet-cell footer-cell" id="media-reactiva">-</div>
            <div class="sheet-cell footer-cell"></div>
        `;
        return html + `</div>`;
    },

    renderEdarBotosHoja2Content() {
        const logs = JSON.parse(localStorage.getItem(`logs_EDAR_BOTOS`) || '[]');
        const equipment = [
            { id: 'b1', label: 'BOMBA Nº 1' }, { id: 'b2', label: 'BOMBA Nº 2' },
            { id: 'achique', label: 'BOMBA ACHIQUE' }, { id: 'masko', label: 'MASKO-ZOLL' },
            { id: 'agitador', label: 'AGITADOR' }, { id: 'parrilla', label: 'PARRILLA' },
            { id: 'barredera', label: 'BARREDERA' }, { id: 'tornillo', label: 'TORNILLO' },
            { id: 'fango', label: 'BOMBA FANGO' }
        ];

        let html = `<div class="sheet-table botos-horas">
            <div class="sheet-cell sheet-header-cell span-row-2"></div>
            <div class="sheet-cell sheet-header-cell span-row-2">FECHA</div>
            <div class="sheet-cell sheet-header-cell span-row-2" style="display:none;">HORA</div>
            <div class="sheet-cell sheet-header-cell span-row-2"></div>`; // Separador vertical
        equipment.forEach(e => { html += `<div class="sheet-cell sheet-header-cell span-col-2">${e.label}</div> `; });
        html += equipment.map(() => `<div class="sheet-cell sheet-header-cell">LECTURA</div><div class="sheet-cell sheet-header-cell">H.</div>`).join('');

        // Blank horizontal row
        for (let i = 0; i < 22; i++) html += `<div class="sheet-cell"></div>`;

        const monthKey = `${this.currentYear}-${String(this.currentMonth + 1).padStart(2, '0')}`;
        const initialLog = logs.find(l => l.fecha === `${monthKey}-00`) || {};

        // Initial Row (Day 00)
        html += `<div class="sheet-cell initial-row-cell">-</div>`;
        html += `<div class="sheet-cell">
            <input type="text" class="row-input" data-date="${monthKey}-00" data-field="h2_fecha_manual" value="${initialLog.h2_fecha_manual || ''}" placeholder="INICIAL" >
        </div> `;
        html += `<div class="sheet-cell" style="display:none;"><input type="time" class="row-input" data-date="${monthKey}-00" data-field="hora" value="${initialLog.hora || ''}" ></div>`;
        html += `<div class="sheet-cell"></div>`; // Separador vertical
        equipment.forEach(e => {
            html += `<div class="sheet-cell"><input type="number" step="0.1" class="row-input" data-date="${monthKey}-00" data-field="h2_${e.id}_l" value="${initialLog['h2_' + e.id + '_l'] || ''}" ></div>`;
            html += `<div class="sheet-cell"><input type="number" step="0.1" class="row-input" data-date="${monthKey}-00" data-field="h2_${e.id}_h" value="${initialLog['h2_' + e.id + '_h'] || ''}" ></div>`;
        });

        const daysInMonth = new Date(this.currentYear, this.currentMonth + 1, 0).getDate();
        for (let d = 1; d <= 31; d++) {
            const dateStr = `${this.currentYear}-${String(this.currentMonth + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
            const log = logs.find(l => l.fecha === dateStr) || {};
            const isInvalidDay = d > daysInMonth;

            html += `<div class="sheet-cell"></div>`;
            html += `<div class="sheet-cell date-cell ${isInvalidDay ? 'disabled-row' : ''}">
                <input type="text" class="row-input" data-date="${dateStr}" data-field="h2_fecha_manual" value="${log.h2_fecha_manual || ''}" placeholder="" ${isInvalidDay ? 'disabled' : ''}>
            </div> `;
            html += `<div class="sheet-cell" style="display:none;"><input type="time" class="row-input" data-date="${dateStr}" data-field="hora" value="${log.hora || ''}" ${isInvalidDay ? 'disabled' : ''}></div>`;
            html += `<div class="sheet-cell"></div>`; // Separador vertical
            equipment.forEach(e => {
                html += `<div class="sheet-cell"><input type="number" step="0.1" class="row-input" data-date="${dateStr}" data-field="h2_${e.id}_l" value="${log['h2_' + e.id + '_l'] || ''}" ${isInvalidDay ? 'disabled' : ''}></div>`;
                html += `<div class="sheet-cell"><input type="number" step="0.1" class="row-input" data-date="${dateStr}" data-field="h2_${e.id}_h" value="${log['h2_' + e.id + '_h'] || ''}" ${isInvalidDay ? 'disabled' : ''}></div>`;
            });
        }

        html += `<div class="sheet-cell footer-cell span-col-4">TOTAL</div>`;
        equipment.forEach(e => {
            html += `<div class="sheet-cell footer-cell" id="total-h2-${e.id}-l">-</div><div class="sheet-cell footer-cell" id="total-h2-${e.id}-h">-</div>`;
        });
        html += `<div class="sheet-cell footer-cell span-col-4">MEDIA</div>`;
        equipment.forEach(e => {
            html += `<div class="sheet-cell footer-cell" id="media-h2-${e.id}-l">-</div><div class="sheet-cell footer-cell" id="media-h2-${e.id}-h">-</div>`;
        });
        return html + `</div>`;
    },

    renderEdarBotosHoja3Content() {
        const logs = JSON.parse(localStorage.getItem(`logs_EDAR_BOTOS`) || '[]');
        const equipment = [
            { id: 'dac1', label: 'BOMBA DAC Nº 1' }, { id: 'dac2', label: 'BOMBA DAC Nº 2' },
            { id: 'comp', label: 'COMPRESOR' }, { id: 'floc1', label: 'FLOCULANTE 1' },
            { id: 'floc2', label: 'FLOCULANTE 2' }, { id: 'coag1', label: 'COAGULANTE 1' },
            { id: 'coag2', label: 'COAGULANTE 2' }, { id: 'sopl1', label: 'SOPLANTE 1' },
            { id: 'sopl2', label: 'SOPLANTE 2' }
        ];

        let html = `<div class="sheet-table botos-horas">
            <div class="sheet-cell sheet-header-cell span-row-2"></div>
            <div class="sheet-cell sheet-header-cell span-row-2">FECHA</div>
            <div class="sheet-cell sheet-header-cell span-row-2" style="display:none;">HORA</div>
            <div class="sheet-cell sheet-header-cell span-row-2"></div>`; // Separador vertical
        equipment.forEach(e => { html += `<div class="sheet-cell sheet-header-cell span-col-2">${e.label}</div> `; });
        html += equipment.map(() => `<div class="sheet-cell sheet-header-cell">LECTURA</div><div class="sheet-cell sheet-header-cell">H.</div>`).join('');

        const monthKey = `${this.currentYear}-${String(this.currentMonth + 1).padStart(2, '0')}`;
        const initialLog = logs.find(l => l.fecha === `${monthKey}-00`) || {};

        // Blank horizontal row
        for (let i = 0; i < 22; i++) html += `<div class="sheet-cell"></div>`;

        // Initial Row (Day 00)
        html += `<div class="sheet-cell initial-row-cell">-</div>`;
        html += `<div class="sheet-cell date-cell">
            <input type="text" class="row-input" data-date="${monthKey}-00" data-field="h3_fecha_manual" value="${initialLog.h3_fecha_manual || ''}" placeholder="INICIAL" >
        </div> `;
        html += `<div class="sheet-cell" style="display:none;"><input type="time" class="row-input" data-date="${monthKey}-00" data-field="hora" value="${initialLog.hora || ''}" ></div>`;
        html += `<div class="sheet-cell"></div>`; // Separador vertical
        equipment.forEach(e => {
            html += `<div class="sheet-cell"><input type="number" step="0.1" class="row-input" data-date="${monthKey}-00" data-field="h3_${e.id}_l" value="${initialLog['h3_' + e.id + '_l'] || ''}" ></div>`;
            html += `<div class="sheet-cell"><input type="number" step="0.1" class="row-input" data-date="${monthKey}-00" data-field="h3_${e.id}_h" value="${initialLog['h3_' + e.id + '_h'] || ''}" ></div>`;
        });

        const daysInMonth = new Date(this.currentYear, this.currentMonth + 1, 0).getDate();
        for (let d = 1; d <= 31; d++) {
            const dateStr = `${this.currentYear}-${String(this.currentMonth + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
            const log = logs.find(l => l.fecha === dateStr) || {};
            const isInvalidDay = d > daysInMonth;

            html += `<div class="sheet-cell"></div>`;
            html += `<div class="sheet-cell date-cell ${isInvalidDay ? 'disabled-row' : ''}">
                <input type="text" class="row-input" data-date="${dateStr}" data-field="h3_fecha_manual" value="${log.h3_fecha_manual || ''}" placeholder="" ${isInvalidDay ? 'disabled' : ''}>
            </div> `;
            html += `<div class="sheet-cell" style="display:none;"><input type="time" class="row-input" data-date="${dateStr}" data-field="hora" value="${log.hora || ''}" ${isInvalidDay ? 'disabled' : ''}></div>`;
            html += `<div class="sheet-cell"></div>`; // Separador vertical
            equipment.forEach(e => {
                html += `<div class="sheet-cell"><input type="number" step="0.1" class="row-input" data-date="${dateStr}" data-field="h3_${e.id}_l" value="${log['h3_' + e.id + '_l'] || ''}" ${isInvalidDay ? 'disabled' : ''}></div>`;
                html += `<div class="sheet-cell"><input type="number" step="0.1" class="row-input" data-date="${dateStr}" data-field="h3_${e.id}_h" value="${log['h3_' + e.id + '_h'] || ''}" ${isInvalidDay ? 'disabled' : ''}></div>`;
            });
        }

        html += `<div class="sheet-cell footer-cell span-col-4">TOTAL</div>`;
        equipment.forEach(e => {
            html += `<div class="sheet-cell footer-cell" id="total-h3-${e.id}-l">-</div><div class="sheet-cell footer-cell" id="total-h3-${e.id}-h">-</div>`;
        });
        html += `<div class="sheet-cell footer-cell span-col-4">MEDIA</div>`;
        equipment.forEach(e => {
            html += `<div class="sheet-cell footer-cell" id="media-h3-${e.id}-l">-</div><div class="sheet-cell footer-cell" id="media-h3-${e.id}-h">-</div>`;
        });
        return html + `</div>`;
    },

    renderEdarBotosHoja4Content() {
        const logs = JSON.parse(localStorage.getItem(`logs_EDAR_BOTOS`) || '[]');
        const equipment = [
            { id: 'mono', label: 'BOMBA MONO' }, { id: 'poli', label: 'POLIELECTROLITO' },
            { id: 'agit', label: 'AGITADOR' }, { id: 'banda', label: 'FILTRO BANDA' },
            { id: 'v1', label: '' }, { id: 'v2', label: '' }, { id: 'v3', label: '' }, { id: 'v4', label: '' }, { id: 'v5', label: '' }
        ];

        let html = `<div class="sheet-table botos-horas">
            <div class="sheet-cell sheet-header-cell span-row-2"></div>
            <div class="sheet-cell sheet-header-cell span-row-2">FECHA</div>
            <div class="sheet-cell sheet-header-cell span-row-2" style="display:none;">HORA</div>
            <div class="sheet-cell sheet-header-cell span-row-2"></div>`; // Separador vertical
        equipment.forEach(e => { html += `<div class="sheet-cell sheet-header-cell span-col-2">${e.label || '-'}</div>`; });
        html += equipment.map(() => `<div class="sheet-cell sheet-header-cell">LECTURA</div><div class="sheet-cell sheet-header-cell">H.</div>`).join('');

        const monthKey = `${this.currentYear}-${String(this.currentMonth + 1).padStart(2, '0')}`;
        const initialLog = logs.find(l => l.fecha === `${monthKey}-00`) || {};

        // Blank horizontal row
        for (let i = 0; i < 22; i++) html += `<div class="sheet-cell"></div>`;

        // Initial Row (Day 00)
        html += `<div class="sheet-cell initial-row-cell">-</div>`;
        html += `<div class="sheet-cell date-cell">
            <input type="text" class="row-input" data-date="${monthKey}-00" data-field="h4_fecha_manual" value="${initialLog.h4_fecha_manual || ''}" placeholder="INICIAL" >
        </div>`;
        html += `<div class="sheet-cell" style="display:none;"><input type="time" class="row-input" data-date="${monthKey}-00" data-field="hora" value="${initialLog.hora || ''}" ></div>`;
        html += `<div class="sheet-cell"></div>`; // Separador vertical
        equipment.forEach(e => {
            html += `<div class="sheet-cell"><input type="number" step="0.1" class="row-input" data-date="${monthKey}-00" data-field="h4_${e.id}_l" value="${initialLog['h4_' + e.id + '_l'] || ''}" ></div>`;
            html += `<div class="sheet-cell"><input type="number" step="0.1" class="row-input" data-date="${monthKey}-00" data-field="h4_${e.id}_h" value="${initialLog['h4_' + e.id + '_h'] || ''}" ></div>`;
        });

        const daysInMonth = new Date(this.currentYear, this.currentMonth + 1, 0).getDate();
        for (let d = 1; d <= 31; d++) {
            const dateStr = `${this.currentYear}-${String(this.currentMonth + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
            const log = logs.find(l => l.fecha === dateStr) || {};
            const isInvalidDay = d > daysInMonth;

            html += `<div class="sheet-cell"></div>`;
            html += `<div class="sheet-cell date-cell ${isInvalidDay ? 'disabled-row' : ''}">
                <input type="text" class="row-input" data-date="${dateStr}" data-field="h4_fecha_manual" value="${log.h4_fecha_manual || ''}" placeholder="" ${isInvalidDay ? 'disabled' : ''}>
            </div>`;
            html += `<div class="sheet-cell" style="display:none;"><input type="time" class="row-input" data-date="${dateStr}" data-field="hora" value="${log.hora || ''}" ${isInvalidDay ? 'disabled' : ''}></div>`;
            html += `<div class="sheet-cell"></div>`; // Separador vertical
            equipment.forEach(e => {
                html += `<div class="sheet-cell"><input type="number" step="0.1" class="row-input" data-date="${dateStr}" data-field="h4_${e.id}_l" value="${log['h4_' + e.id + '_l'] || ''}" ${isInvalidDay ? 'disabled' : ''}></div>`;
                html += `<div class="sheet-cell"><input type="number" step="0.1" class="row-input" data-date="${dateStr}" data-field="h4_${e.id}_h" value="${log['h4_' + e.id + '_h'] || ''}" ${isInvalidDay ? 'disabled' : ''}></div>`;
            });
        }

        html += `<div class="sheet-cell footer-cell span-col-4">TOTAL</div>`;
        equipment.forEach(e => {
            html += `<div class="sheet-cell footer-cell" id="total-h4-${e.id}-l">-</div><div class="sheet-cell footer-cell" id="total-h4-${e.id}-h">-</div>`;
        });
        html += `<div class="sheet-cell footer-cell span-col-4">MEDIA</div>`;
        equipment.forEach(e => {
            html += `<div class="sheet-cell footer-cell" id="media-h4-${e.id}-l">-</div><div class="sheet-cell footer-cell" id="media-h4-${e.id}-h">-</div>`;
        });
        return html + `</div>`;
    },

    renderEdarBotosEnergiaContent() {
        const logs = JSON.parse(localStorage.getItem(`logs_EDAR_BOTOS`) || '[]');
        return this.renderGenericEnergiaContent(logs, 'EDAR_BOTOS');
    },

    renderGenericEnergiaContent(logs, station) {
        const monthKey = `${this.currentYear}-${String(this.currentMonth + 1).padStart(2, '0')}`;
        let html = `<div class="sheet-table botos-energia">
            <!-- HEADER ROW 1: Sized for 21 columns total (1 DIA + 20 data) -->
            <div class="sheet-cell sheet-header-cell span-row-2" style="background:#e3f2fd; border-bottom:2px solid #000;">DIA</div>
            <div class="sheet-cell sheet-header-cell span-col-8" style="background:#fff3e0;">ENERGÍA ACTIVA (kWh)</div>
            <div class="sheet-cell sheet-header-cell span-col-8" style="background:#f1f8e9;">ENERGÍA REACTIVA (kVArh)</div>
            <div class="sheet-cell sheet-header-cell span-col-4" style="background:#fce4ec;">MAXÍMETRO (kW)</div>

            <!-- HEADER ROW 2: Sub-labels -->
            <div class="sheet-cell sheet-header-cell">P1 1.18.1</div><div class="sheet-cell sheet-header-cell">P2 1.18.2</div>
            <div class="sheet-cell sheet-header-cell">P3 1.18.3</div><div class="sheet-cell sheet-header-cell">P4 1.18.4</div>
            <div class="sheet-cell sheet-header-cell">P5 1.18.5</div><div class="sheet-cell sheet-header-cell">P6 1.18.6</div>
            <div class="sheet-cell sheet-header-cell" style="color:#e65100;">TOTAL 1.18.0</div><div class="sheet-cell sheet-header-cell" style="background:#eee;">DIF</div>
            
            <div class="sheet-cell sheet-header-cell">P1 1.58.1</div><div class="sheet-cell sheet-header-cell">P2 1.58.2</div>
            <div class="sheet-cell sheet-header-cell">P3 1.58.3</div><div class="sheet-cell sheet-header-cell">P4 1.58.4</div>
            <div class="sheet-cell sheet-header-cell">P5 1.58.5</div><div class="sheet-cell sheet-header-cell">P6 1.58.6</div>
            <div class="sheet-cell sheet-header-cell" style="color:#2e7d32;">TOTAL 1.58.0</div><div class="sheet-cell sheet-header-cell" style="background:#eee;">DIF</div>

            <div class="sheet-cell sheet-header-cell">P.1.16.1</div><div class="sheet-cell sheet-header-cell">P.1.16.2</div>
            <div class="sheet-cell sheet-header-cell">P.1.16.3</div><div class="sheet-cell sheet-header-cell" style="color:#c2185b;">TOTAL 1.16.0</div>
        `;

        const daysInMonth = new Date(this.currentYear, this.currentMonth + 1, 0).getDate();
        for (let d = 0; d <= 31; d++) {
            const dateStr = `${monthKey}-${String(d).padStart(2, '0')}`;
            const log = logs.find(l => l.fecha === dateStr) || {};
            const isDisabled = (d > daysInMonth && d !== 0);
            const isInitial = (d === 0);
            const rowStyle = isInitial ? 'style="background:#fff9c4; font-weight:bold;"' : '';

            html += `<div class="sheet-cell ${isDisabled ? 'disabled-day' : ''}" ${rowStyle}>${d}</div>`;

            // ACTIVA (8 cells: 6 periods + Total + Dif)
            ['e_a_p1', 'e_a_p2', 'e_a_p3', 'e_a_p4', 'e_a_p5', 'e_a_p6', 'e_a_total'].forEach(f => {
                html += `<div class="sheet-cell ${isDisabled ? 'disabled-day' : ''}"><input type="number" step="0.001" class="row-input" data-date="${dateStr}" data-field="${f}" value="${log[f] || ''}" ${isDisabled ? 'disabled' : ''}></div>`;
            });
            html += `<div class="sheet-cell" style="background:#fafafa;"><input type="number" class="row-input" data-date="${dateStr}" data-field="e_a_dif" value="${log.e_a_dif || ''}" disabled></div>`;

            // REACTIVA (8 cells: 6 periods + Total + Dif)
            ['e_r_p1', 'e_r_p2', 'e_r_p3', 'e_r_p4', 'e_r_p5', 'e_r_p6', 'e_r_total'].forEach(f => {
                html += `<div class="sheet-cell ${isDisabled ? 'disabled-day' : ''}"><input type="number" step="0.001" class="row-input" data-date="${dateStr}" data-field="${f}" value="${log[f] || ''}" ${isDisabled ? 'disabled' : ''}></div>`;
            });
            html += `<div class="sheet-cell" style="background:#fafafa;"><input type="number" class="row-input" data-date="${dateStr}" data-field="e_r_dif" value="${log.e_r_dif || ''}" disabled></div>`;

            // MAXIMETRO (4 cells: 3 periods + Total)
            ['m_p1', 'm_p2', 'm_p3', 'm_total'].forEach(f => {
                html += `<div class="sheet-cell ${isDisabled ? 'disabled-day' : ''}"><input type="number" step="0.001" class="row-input" data-date="${dateStr}" data-field="${f}" value="${log[f] || ''}" ${isDisabled ? 'disabled' : ''}></div>`;
            });
        }
        return html + `</div>`;
    },

    setupMonthlyEventListeners() {
        const monthSelect = document.getElementById('month-select');
        if (monthSelect) {
            monthSelect.onchange = (e) => {
                this.currentMonth = parseInt(e.target.value);
                this.renderFormFields(this.currentStation);
            };
        }
        const yearSelect = document.getElementById('year-select');
        if (yearSelect) {
            yearSelect.onchange = (e) => {
                this.currentYear = parseInt(e.target.value);
                this.renderFormFields(this.currentStation);
            };
        }

        document.querySelectorAll('.row-input').forEach(input => {
            input.onchange = () => {
                const date = input.getAttribute('data-date');
                const field = input.getAttribute('data-field');
                const val = input.value;
                this.saveSingleField(date, field, val);

                // Auto-save hora when any field in that row is edited (only if not already set)
                if (field !== 'hora') {
                    const horaInput = document.querySelector(`.row-input[data-date="${date}"][data-field="hora"]`);
                    if (horaInput && !horaInput.getAttribute('data-hora-set')) {
                        const autoHora = new Date().toTimeString().substring(0, 5);
                        horaInput.value = autoHora;
                        horaInput.setAttribute('data-hora-set', 'true');
                        this.saveSingleField(date, 'hora', autoHora);
                    }
                }

                if (field === 'caudal_lect') this.recalculateDailyConsumption();

                // If in an energy sheet, recalculate the DIF automatically
                if (this.currentBotosTab === 'hoja_energia' || this.currentCorredoiraTab === 'hoja_energia' || this.currentVilatuxeTab === 'hoja2') {
                    if (field === 'e_a_total' || field === 'e_r_total') {
                        this.recalculateEnergiaDifs(this.currentStation === 'EDAR_BOTOS' ? 'EDAR_BOTOS' : (this.currentStation === 'EDAR_CORREDOIRA' ? 'EDAR_CORREDOIRA' : 'VILATUXE'));
                    }
                }

                if (['BOMBEO_BOTOS', 'CATASOS', 'VILATUXE'].includes(this.currentStation) && ['b1', 'b2', 'b3'].includes(field)) {
                    if (this.currentStation === 'BOMBEO_BOTOS') this.recalculateBombeoBotosDifs();
                    if (this.currentStation === 'CATASOS') this.recalculateCatasosDifs();
                }
                this.recalculateTotals();
            };
        });
    },

    recalculateDailyConsumption() {
        const daysInMonth = new Date(this.currentYear, this.currentMonth + 1, 0).getDate();
        const monthKey = `${this.currentYear}-${String(this.currentMonth + 1).padStart(2, '0')}`;

        // Iterate days 1 to 31
        for (let d = 1; d <= 31; d++) {
            const currentDayKey = `${monthKey}-${String(d).padStart(2, '0')}`;
            const prevDayKey = `${monthKey}-${String(d - 1).padStart(2, '0')}`;

            const currentLectInput = document.querySelector(`.row-input[data-date="${currentDayKey}"][data-field="caudal_lect"]`);
            const prevLectInput = document.querySelector(`.row-input[data-date="${prevDayKey}"][data-field="caudal_lect"]`);
            const targetM3dInput = document.querySelector(`.row-input[data-date="${currentDayKey}"][data-field="caudal_m3d"]`);

            if (currentLectInput && prevLectInput && targetM3dInput) {
                const currentVal = parseFloat(currentLectInput.value);
                const prevVal = parseFloat(prevLectInput.value);

                if (!isNaN(currentVal) && !isNaN(prevVal)) {
                    targetM3dInput.value = (currentVal - prevVal).toFixed(1);
                }
            }
        }
    },

    recalculateBombeoBotosDifs() {
        const monthKey = `${this.currentYear}-${String(this.currentMonth + 1).padStart(2, '0')}`;
        for (let d = 1; d <= 31; d++) {
            const currentDayStr = `${monthKey}-${String(d).padStart(2, '0')}`;
            const prevDayStr = `${monthKey}-${String(d - 1).padStart(2, '0')}`;
            ['b1', 'b2', 'b3'].forEach(bomb => {
                const currentInput = document.querySelector(`.row-input[data-date="${currentDayStr}"][data-field="${bomb}"]`);
                const prevInput = document.querySelector(`.row-input[data-date="${prevDayStr}"][data-field="${bomb}"]`);
                const diffInput = document.querySelector(`.row-input[data-date="${currentDayStr}"][data-field="${bomb}_dif"]`);
                if (currentInput && prevInput && diffInput) {
                    const currentVal = parseFloat(currentInput.value);
                    const prevVal = parseFloat(prevInput.value);
                    if (!isNaN(currentVal) && !isNaN(prevVal)) {
                        diffInput.value = (currentVal - prevVal).toFixed(0);
                    }
                }
            });
        }
    },

    recalculateCatasosDifs() {
        const monthKey = `${this.currentYear}-${String(this.currentMonth + 1).padStart(2, '0')}`;
        for (let d = 1; d <= 31; d++) {
            const currentDayStr = `${monthKey}-${String(d).padStart(2, '0')}`;
            const prevDayStr = `${monthKey}-${String(d - 1).padStart(2, '0')}`;
            ['b1', 'b2', 'b3'].forEach(bomb => {
                const currentInput = document.querySelector(`.row-input[data-date="${currentDayStr}"][data-field="${bomb}"]`);
                const prevInput = document.querySelector(`.row-input[data-date="${prevDayStr}"][data-field="${bomb}"]`);
                const diffInput = document.querySelector(`.row-input[data-date="${currentDayStr}"][data-field="${bomb}_dif"]`);
                if (currentInput && prevInput && diffInput) {
                    const currentVal = parseFloat(currentInput.value);
                    const prevVal = parseFloat(prevInput.value);
                    if (!isNaN(currentVal) && !isNaN(prevVal)) {
                        diffInput.value = (currentVal - prevVal).toFixed(0);
                    }
                }
            });
        }
    },

    recalculateEdarBotosEnergiaDifs() {
        const monthKey = `${this.currentYear}-${String(this.currentMonth + 1).padStart(2, '0')}`;
        for (let d = 1; d <= 31; d++) {
            const currentDayStr = `${monthKey}-${String(d).padStart(2, '0')}`;
            const prevDayStr = `${monthKey}-${String(d - 1).padStart(2, '0')}`;

            // Energía Activa DIF
            const currentActTotal = document.querySelector(`.row-input[data-date="${currentDayStr}"][data-field="e_a_total"]`);
            const prevActTotal = document.querySelector(`.row-input[data-date="${prevDayStr}"][data-field="e_a_total"]`);
            const diffActInput = document.querySelector(`.row-input[data-date="${currentDayStr}"][data-field="e_a_dif"]`);

            if (currentActTotal && prevActTotal && diffActInput) {
                const currentVal = parseFloat(currentActTotal.value);
                const prevVal = parseFloat(prevActTotal.value);
                if (!isNaN(currentVal) && !isNaN(prevVal)) {
                    const diff = (currentVal - prevVal).toFixed(0);
                    diffActInput.value = diff;
                }
            }

            // Energía Reactiva DIF
            const currentReacTotal = document.querySelector(`.row-input[data-date="${currentDayStr}"][data-field="e_r_total"]`);
            const prevReacTotal = document.querySelector(`.row-input[data-date="${prevDayStr}"][data-field="e_r_total"]`);
            const diffReacInput = document.querySelector(`.row-input[data-date="${currentDayStr}"][data-field="e_r_dif"]`);

            if (currentReacTotal && prevReacTotal && diffReacInput) {
                const currentVal = parseFloat(currentReacTotal.value);
                const prevVal = parseFloat(prevReacTotal.value);
                if (!isNaN(currentVal) && !isNaN(prevVal)) {
                    const diff = (currentVal - prevVal).toFixed(0);
                    diffReacInput.value = diff;
                }
            }
        }
    },

    recalculateTotals() {
        if (this.currentStation === 'ETAP') {
            this.recalculateEtapTotals();
        } else if (this.currentStation === 'BOMBEO_BOTOS') {
            this.recalculateBombeoBotosDifs();
        } else if (this.currentStation === 'CATASOS') {
            this.recalculateCatasosDifs();
        } else if (this.currentStation === 'EDAR_BOTOS') {
            if (this.currentBotosTab === 'hoja_energia') {
                this.recalculateEdarBotosEnergiaDifs();
            }
            if (this.currentBotosTab === 'hoja1') {
                const fields = ['caudal_lect', 'caudal_m3d', 'desbaste', 'fangos_filtro', 'pax', 'polielectrolito', 'punta', 'valle', 'llano', 'reactiva'];
                const idMap = {
                    'caudal_lect': 'caudal-lect', 'caudal_m3d': 'caudal-m3d', 'desbaste': 'desbaste',
                    'fangos_filtro': 'fangos', 'pax': 'pax', 'polielectrolito': 'poli', 'punta': 'punta',
                    'valle': 'valle', 'llano': 'llano', 'reactiva': 'reactiva'
                };
                this.calculateFields(fields, idMap);
            } else {
                // Hojas 2, 3, 4 - Horizontal equipment totals
                const tabNum = this.currentBotosTab.replace('hoja', '');
                // Identify all equipment in the current tab
                const equipmentInputs = document.querySelectorAll(`.row-input[data-field^="h${tabNum}_"]`);
                const equipmentIds = new Set();
                equipmentInputs.forEach(input => {
                    const field = input.getAttribute('data-field'); // e.g. h2_b1_l
                    const parts = field.split('_');
                    if (parts.length >= 2) equipmentIds.add(parts[1]);
                });

                equipmentIds.forEach(eid => {
                    ['l', 'h'].forEach(suffix => {
                        const field = `h${tabNum}_${eid}_${suffix}`;
                        const inputs = document.querySelectorAll(`.row-input[data-field="${field}"]:not([disabled])`);
                        let total = 0;
                        let count = 0;
                        inputs.forEach(input => {
                            const val = parseFloat(input.value);
                            if (!isNaN(val)) { total += val; count++; }
                        });
                        const totalEl = document.getElementById(`total-h${tabNum}-${eid}-${suffix}`);
                        const mediaEl = document.getElementById(`media-h${tabNum}-${eid}-${suffix}`);
                        if (totalEl) totalEl.textContent = total > 0 ? total.toFixed(1) : '-';
                        if (mediaEl) mediaEl.textContent = (count > 0 && total > 0) ? (total / count).toFixed(2) : '-';
                    });
                });
            }
        } else if (this.currentStation === 'EDAR_CORREDOIRA') {
            if (this.currentCorredoiraTab === 'diario') {
                const fields = ['precipitacion', 'caudal_m3d', 'res_desbaste', 'res_arenas', 'res_fangos', 'cons_punta', 'cons_valle', 'cons_llano', 'reactiva', 'taire', 'tagua', 'caudal_lect', 'oxigeno', 'vol_fango'];
                const idMap = {
                    'precipitacion': 'corredoira-precipitacion', 'caudal_m3d': 'corredoira-caudal',
                    'res_desbaste': 'corredoira-desbaste', 'res_arenas': 'corredoira-arenas',
                    'res_fangos': 'corredoira-fangos', 'cons_punta': 'corredoira-punta',
                    'cons_valle': 'corredoira-valle', 'cons_llano': 'corredoira-llano',
                    'reactiva': 'corredoira-reactiva', 'taire': 'corredoira-taire',
                    'tagua': 'corredoira-tagua', 'caudal_lect': 'corredoira-caudal-lect',
                    'oxigeno': 'corredoira-oxigeno', 'vol_fango': 'corredoira-volfang'
                };
                this.calculateFields(fields, idMap);
            } else {
                const tabNum = this.currentCorredoiraTab.replace('hoja', '');
                const equipmentInputs = document.querySelectorAll(`.row-input[data-field^="h${tabNum}_"]`);
                const equipmentIds = new Set();
                equipmentInputs.forEach(input => {
                    const field = input.getAttribute('data-field');
                    const parts = field.split('_');
                    if (parts.length >= 2) equipmentIds.add(parts[1]);
                });

                equipmentIds.forEach(eid => {
                    ['l', 'h'].forEach(suffix => {
                        const field = `h${tabNum}_${eid}_${suffix}`;
                        const inputs = document.querySelectorAll(`.row-input[data-field="${field}"]:not([disabled])`);
                        let total = 0;
                        let count = 0;
                        inputs.forEach(input => {
                            const val = parseFloat(input.value);
                            if (!isNaN(val)) { total += val; count++; }
                        });
                        const totalEl = document.getElementById(`total-h${tabNum}-${eid}-${suffix}`);
                        const mediaEl = document.getElementById(`media-h${tabNum}-${eid}-${suffix}`);
                        if (totalEl) totalEl.textContent = total > 0 ? total.toFixed(1) : '-';
                        if (mediaEl) mediaEl.textContent = (count > 0 && total > 0) ? (total / count).toFixed(2) : '-';
                    });
                });
            }
        }
    },

    calculateFields(fields, idMap) {
        fields.forEach(field => {
            const inputs = document.querySelectorAll(`.row-input[data-field="${field}"]:not([disabled])`);
            let total = 0;
            let count = 0;
            inputs.forEach(input => {
                const val = parseFloat(input.value);
                if (!isNaN(val)) {
                    total += val;
                    count++;
                }
            });

            const totalEl = document.getElementById(`total-${idMap[field]}`);
            const mediaEl = document.getElementById(`media-${idMap[field]}`);
            if (totalEl) totalEl.textContent = total > 0 ? total.toFixed(1) : '-';
            if (mediaEl) mediaEl.textContent = (count > 0 && total > 0) ? (total / count).toFixed(2) : '-';
        });
    },

    renderEdarCorredoiraSheet(container) {
        if (!this.currentMonth) {
            const now = new Date();
            this.currentMonth = now.getMonth();
            this.currentYear = now.getFullYear();
        }


        const months = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"];

        let html = `
            <div class="premium-header">
                <div class="header-logo-top">AQUADEZA</div>
                <div class="header-main-titles">
                    <div class="brand-tagline">SERVICIO DE LALÍN</div>
                    <h1 class="station-title">EDAR CORREDOIRA</h1>
                    <div class="sheet-info-badge">${this.currentCorredoiraTab.toUpperCase()}</div>
                </div>
                <div class="header-controls-row">
                    <div class="control-pill">
                        <label>MES</label>
                        <select id="month-select">
                            ${months.map((m, i) => `<option value="${i}" ${i === this.currentMonth ? 'selected' : ''}>${m}</option>`).join('')}
                        </select>
                    </div>
                    <div class="control-pill">
                        <label>AÑO</label>
                        <select id="year-select">
                            ${this.generateYearOptions(this.currentYear)}
                        </select>
                    </div>
                </div>
            </div>
            
            <div class="sheet-tabs">
                <button class="tab-btn ${this.currentCorredoiraTab === 'diario' ? 'active' : ''}" data-tab="diario">DIARIO</button>
                <button class="tab-btn ${this.currentCorredoiraTab === 'hoja_energia' ? 'active' : ''}" data-tab="hoja_energia">CUADRO ENERGÍA</button>
                <button class="tab-btn ${this.currentCorredoiraTab === 'hoja1' ? 'active' : ''}" data-tab="hoja1">HORAS 1</button>
                <button class="tab-btn ${this.currentCorredoiraTab === 'hoja2' ? 'active' : ''}" data-tab="hoja2">HORAS 2</button>
                <button class="tab-btn ${this.currentCorredoiraTab === 'hoja3' ? 'active' : ''}" data-tab="hoja3">HORAS 3</button>
            </div>
        `;

        if (this.currentCorredoiraTab === 'diario') {
            html += this.renderCorredoiraDiarioContent();
        } else if (this.currentCorredoiraTab === 'hoja_energia') {
            html += this.renderEdarCorredoiraEnergiaContent();
        } else if (this.currentCorredoiraTab === 'hoja1') {
            html += this.renderCorredoiraHoja1Content();
        } else if (this.currentCorredoiraTab === 'hoja2') {
            html += this.renderCorredoiraHoja2Content();
        } else if (this.currentCorredoiraTab === 'hoja3') {
            html += this.renderCorredoiraHoja3Content();
        }

        html += `
            <div class="monthly-footer-controls" style="padding: 1rem; text-align: right;">
                <button type="submit" class="btn-primary" style="width: auto; padding: 10px 30px;">GUARDAR TODO EL MES</button>
            </div>
        `;

        container.innerHTML = html;

        this.setupMonthlyEventListeners();
        this.setupCorredoiraTabListeners();
        this.recalculateTotals();
    },

    setupCorredoiraTabListeners() {
        document.querySelectorAll('.tab-btn').forEach(btn => {
            btn.onclick = () => {
                this.currentCorredoiraTab = btn.dataset.tab;
                this.renderEdarCorredoiraSheet(document.getElementById('form-fields'));
            };
        });
    },

    renderCorredoiraDiarioContent() {
        let html = `
        <div class="sheet-table corredoira">
            <!-- Row 3: Grouped Headers -->
            <div class="sheet-cell sheet-header-cell span-row-2"><div class="header-horizontal">Día</div></div>
            <div class="sheet-cell sheet-header-cell span-row-2"><div class="header-horizontal">Hora</div></div>
            <div class="sheet-cell sheet-header-cell span-col-2">Tª</div>
            <div class="sheet-cell sheet-header-cell span-row-2"><div class="header-vertical">Precipitación mm.</div></div>
            <div class="sheet-cell sheet-header-cell span-col-2">CAUDAL</div>
            <div class="sheet-cell sheet-header-cell span-col-2">PARÁMETROS CONTROL</div>
            <div class="sheet-cell sheet-header-cell span-col-3">RESIDUOS</div>
            <div class="sheet-cell sheet-header-cell span-col-4">CONSUMO ELÉCTRICO</div>
            <div class="sheet-cell sheet-header-cell span-row-2"><div class="header-horizontal">OBSERVACIONES</div></div>

            <!-- Row 4: Sub-Headers Row -->
            <div class="sheet-cell sheet-header-cell"><div class="header-vertical">Aire ºC</div></div>
            <div class="sheet-cell sheet-header-cell"><div class="header-vertical">Agua ºC</div></div>
            <div class="sheet-cell sheet-header-cell"><div class="header-horizontal">Lectura</div></div>
            <div class="sheet-cell sheet-header-cell"><div class="header-horizontal">m³/día</div></div>
            <div class="sheet-cell sheet-header-cell"><div class="header-vertical">Oxíg. mg/l</div></div>
            <div class="sheet-cell sheet-header-cell"><div class="header-vertical">Volum. Fang. ml/l</div></div>
            <div class="sheet-cell sheet-header-cell"><div class="header-vertical">Desbaste m³</div></div>
            <div class="sheet-cell sheet-header-cell"><div class="header-vertical">Arenas m³</div></div>
            <div class="sheet-cell sheet-header-cell"><div class="header-vertical">Fangos filtro m³</div></div>
            <div class="sheet-cell sheet-header-cell"><div class="header-horizontal">PUNTA I</div></div>
            <div class="sheet-cell sheet-header-cell"><div class="header-horizontal">VALLE II</div></div>
            <div class="sheet-cell sheet-header-cell"><div class="header-horizontal">LLANO III</div></div>
            <div class="sheet-cell sheet-header-cell"><div class="header-horizontal">ENERGÍA REACTIVA</div></div>
        `;

        const logs = JSON.parse(localStorage.getItem(`logs_EDAR_CORREDOIRA`) || '[]');
        const monthKey = `${this.currentYear}-${String(this.currentMonth + 1).padStart(2, '0')}`;
        const initialLog = logs.find(l => l.fecha === `${monthKey}-00`) || {};

        // 1. Initial Blank Line (Row 1)
        html += `
            <div class="sheet-cell initial-row-cell"></div>
            <div class="sheet-cell" style="display:none;"><input type="time" class="row-input" data-date="${monthKey}-00" data-field="hora" value="${initialLog.hora || ''}" ></div>
            <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${monthKey}-00" data-field="taire" value="${initialLog.taire || ''}" ></div>
            <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${monthKey}-00" data-field="tagua" value="${initialLog.tagua || ''}" ></div>
            <div class="sheet-cell"><input type="text" inputmode="text" class="row-input" data-date="${monthKey}-00" data-field="precipitacion" value="${initialLog.precipitacion || ''}" ></div>
            <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${monthKey}-00" data-field="caudal_lect" value="${initialLog.caudal_lect || ''}" ></div>
            <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${monthKey}-00" data-field="caudal_m3d" value="${initialLog.caudal_m3d || ''}" ></div>
            <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${monthKey}-00" data-field="oxigeno" value="${initialLog.oxigeno || ''}" ></div>
            <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${monthKey}-00" data-field="vol_fango" value="${initialLog.vol_fango || ''}" ></div>
            <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${monthKey}-00" data-field="res_desbaste" value="${initialLog.res_desbaste || ''}" ></div>
            <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${monthKey}-00" data-field="res_arenas" value="${initialLog.res_arenas || ''}" ></div>
            <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${monthKey}-00" data-field="res_fangos" value="${initialLog.res_fangos || ''}" ></div>
            <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${monthKey}-00" data-field="cons_punta" value="${initialLog.cons_punta || ''}" ></div>
            <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${monthKey}-00" data-field="cons_valle" value="${initialLog.cons_valle || ''}" ></div>
            <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${monthKey}-00" data-field="cons_llano" value="${initialLog.cons_llano || ''}" ></div>
            <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${monthKey}-00" data-field="reactiva" value="${initialLog.reactiva || ''}" ></div>
            <div class="sheet-cell"><input type="text" class="row-input" data-date="${monthKey}-00" data-field="observaciones" value="${initialLog.observaciones || ''}" ></div>
        `;

        const daysInMonth = new Date(this.currentYear, this.currentMonth + 1, 0).getDate();
        for (let d = 1; d <= 31; d++) {
            const dateStr = `${this.currentYear}-${String(this.currentMonth + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
            const log = logs.find(l => l.fecha === dateStr) || {};
            const isInvalidDay = d > daysInMonth;

            html += `
                <div class="sheet-cell date-cell ${isInvalidDay ? 'disabled-day' : ''}">${d}</div>
                <div class="sheet-cell" style="display:none;"><input type="time" class="row-input" data-date="${dateStr}" data-field="hora" value="${log.hora || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${dateStr}" data-field="taire" value="${log.taire || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${dateStr}" data-field="tagua" value="${log.tagua || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="text" inputmode="text" class="row-input" data-date="${dateStr}" data-field="precipitacion" value="${log.precipitacion || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${dateStr}" data-field="caudal_lect" value="${log.caudal_lect || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${dateStr}" data-field="caudal_m3d" value="${log.caudal_m3d || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${dateStr}" data-field="oxigeno" value="${log.oxigeno || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${dateStr}" data-field="vol_fango" value="${log.vol_fango || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${dateStr}" data-field="res_desbaste" value="${log.res_desbaste || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${dateStr}" data-field="res_arenas" value="${log.res_arenas || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${dateStr}" data-field="res_fangos" value="${log.res_fangos || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${dateStr}" data-field="cons_punta" value="${log.cons_punta || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${dateStr}" data-field="cons_valle" value="${log.cons_valle || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${dateStr}" data-field="cons_llano" value="${log.cons_llano || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" step="0.1" inputmode="decimal" class="row-input" data-date="${dateStr}" data-field="reactiva" value="${log.reactiva || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="text" class="row-input" data-date="${dateStr}" data-field="observaciones" value="${log.observaciones || ''}" ${isInvalidDay ? 'disabled' : ''}></div>
            `;
        }

        html += `
            <div class="sheet-cell footer-cell">TOTAL</div>
            <div class="sheet-cell footer-cell"></div>
            <div class="sheet-cell footer-cell"></div>
            <div class="sheet-cell footer-cell"></div>
            <div class="sheet-cell footer-cell" id="total-corredoira-precipitacion">-</div>
            <div class="sheet-cell footer-cell"></div>
            <div class="sheet-cell footer-cell" id="total-corredoira-caudal">-</div>
            <div class="sheet-cell footer-cell"></div>
            <div class="sheet-cell footer-cell"></div>
            <div class="sheet-cell footer-cell" id="total-corredoira-desbaste">-</div>
            <div class="sheet-cell footer-cell" id="total-corredoira-arenas">-</div>
            <div class="sheet-cell footer-cell" id="total-corredoira-fangos">-</div>
            <div class="sheet-cell footer-cell" id="total-corredoira-punta">-</div>
            <div class="sheet-cell footer-cell" id="total-corredoira-valle">-</div>
            <div class="sheet-cell footer-cell" id="total-corredoira-llano">-</div>
            <div class="sheet-cell footer-cell" id="total-corredoira-reactiva">-</div>
            <div class="sheet-cell footer-cell"></div>

            <div class="sheet-cell footer-cell">MEDIA</div>
            <div class="sheet-cell footer-cell"></div>
            <div class="sheet-cell footer-cell" id="media-corredoira-taire">-</div>
            <div class="sheet-cell footer-cell" id="media-corredoira-tagua">-</div>
            <div class="sheet-cell footer-cell"></div>
            <div class="sheet-cell footer-cell"></div>
            <div class="sheet-cell footer-cell" id="media-corredoira-caudal">-</div>
            <div class="sheet-cell footer-cell" id="media-corredoira-oxigeno">-</div>
            <div class="sheet-cell footer-cell" id="media-corredoira-volfang">-</div>
            <div class="sheet-cell footer-cell"></div>
            <div class="sheet-cell footer-cell"></div>
            <div class="sheet-cell footer-cell"></div>
            <div class="sheet-cell footer-cell"></div>
            <div class="sheet-cell footer-cell"></div>
            <div class="sheet-cell footer-cell"></div>
            <div class="sheet-cell footer-cell"></div>
            <div class="sheet-cell footer-cell"></div>
        </div>
        `;
        return html;
    },

    renderCorredoiraHoja1Content() {
        const logs = JSON.parse(localStorage.getItem(`logs_EDAR_CORREDOIRA`) || '[]');
        const equipment = [
            { id: 'cinta', label: 'CINTA TRANSP.' }, { id: 'desar', label: 'DESARENADOR' },
            { id: 'barenas', label: 'BOMBA ARENAS' }, { id: 'aerad', label: 'AERADOR' },
            { id: 'careas', label: 'CLASIFICADOR ARENAS' }, { id: 'vaciados', label: 'BOMBA VACIADOS' },
            { id: 'espesador', label: 'ESPESADOR' }, { id: 'filtro', label: 'FILTRO BANDA' }
        ];

        let html = `<div class="sheet-table corredoira-horas">
            <div class="sheet-cell sheet-header-cell span-row-2">DÍA</div>
            <div class="sheet-cell sheet-header-cell span-row-2" style="display:none;">HORA</div>`;
        equipment.forEach(e => { html += `<div class="sheet-cell sheet-header-cell span-col-2">${e.label}</div> `; });
        html += equipment.map(() => `<div class="sheet-cell sheet-header-cell">LECTURA</div><div class="sheet-cell sheet-header-cell">H.</div>`).join('');

        const monthKey = `${this.currentYear}-${String(this.currentMonth + 1).padStart(2, '0')}`;
        const initialLog = logs.find(l => l.fecha === `${monthKey}-00`) || {};

        // Initial Row (Day 00)
        html += `<div class="sheet-cell initial-row-cell">0</div>`;
        html += `<div class="sheet-cell" style="display:none;"><input type="time" class="row-input" data-date="${monthKey}-00" data-field="hora" value="${initialLog.hora || ''}" ></div>`;
        equipment.forEach(e => {
            html += `<div class="sheet-cell"><input type="number" step="0.1" class="row-input" data-date="${monthKey}-00" data-field="h1_${e.id}_l" value="${initialLog['h1_' + e.id + '_l'] || ''}" ></div>`;
            html += `<div class="sheet-cell"><input type="number" step="0.1" class="row-input" data-date="${monthKey}-00" data-field="h1_${e.id}_h" value="${initialLog['h1_' + e.id + '_h'] || ''}" ></div>`;
        });

        const daysInMonth = new Date(this.currentYear, this.currentMonth + 1, 0).getDate();
        for (let d = 1; d <= 31; d++) {
            const dateStr = `${this.currentYear}-${String(this.currentMonth + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
            const log = logs.find(l => l.fecha === dateStr) || {};
            const isInvalidDay = d > daysInMonth;

            html += `<div class="sheet-cell date-cell ${isInvalidDay ? 'disabled-row' : ''}">${d}</div>`;
            html += `<div class="sheet-cell" style="display:none;"><input type="time" class="row-input" data-date="${dateStr}" data-field="hora" value="${log.hora || ''}" ${isInvalidDay ? 'disabled' : ''}></div>`;
            equipment.forEach(e => {
                html += `<div class="sheet-cell"><input type="number" step="0.1" class="row-input" data-date="${dateStr}" data-field="h1_${e.id}_l" value="${log['h1_' + e.id + '_l'] || ''}" ${isInvalidDay ? 'disabled' : ''}></div>`;
                html += `<div class="sheet-cell"><input type="number" step="0.1" class="row-input" data-date="${dateStr}" data-field="h1_${e.id}_h" value="${log['h1_' + e.id + '_h'] || ''}" ${isInvalidDay ? 'disabled' : ''}></div>`;
            });
        }

        html += `<div class="sheet-cell footer-cell span-col-2">TOTAL</div>`;
        equipment.forEach(e => {
            html += `<div class="sheet-cell footer-cell" id="total-h1-${e.id}-l">-</div><div class="sheet-cell footer-cell" id="total-h1-${e.id}-h">-</div>`;
        });
        html += `<div class="sheet-cell footer-cell span-col-2">MEDIA</div>`;
        equipment.forEach(e => {
            html += `<div class="sheet-cell footer-cell" id="media-h1-${e.id}-l">-</div><div class="sheet-cell footer-cell" id="media-h1-${e.id}-h">-</div>`;
        });
        return html + `</div>`;
    },

    renderCorredoiraHoja2Content() {
        const logs = JSON.parse(localStorage.getItem(`logs_EDAR_CORREDOIRA`) || '[]');
        const equipment = [
            { id: 'turba', label: 'TURBINA A' }, { id: 'turbb', label: 'TURBINA B' },
            { id: 'turbc', label: 'TURBINA C' }, { id: 'turbd', label: 'TURBINA D' },
            { id: 'turbe', label: 'TURBINA E' }, { id: 'turbf', label: 'TURBINA F' },
            { id: 'turbg', label: 'TURBINA G' }
        ];

        let html = `<div class="sheet-table corredoira-horas hoja2">
            <div class="sheet-cell sheet-header-cell span-row-2">DÍA</div>
            <div class="sheet-cell sheet-header-cell span-row-2" style="display:none;">HORA</div>`;
        equipment.forEach(e => { html += `<div class="sheet-cell sheet-header-cell span-col-2">${e.label}</div> `; });
        html += equipment.map(() => `<div class="sheet-cell sheet-header-cell">LECTURA</div><div class="sheet-cell sheet-header-cell">H.</div>`).join('');

        const monthKey = `${this.currentYear}-${String(this.currentMonth + 1).padStart(2, '0')}`;
        const initialLog = logs.find(l => l.fecha === `${monthKey}-00`) || {};

        // Initial Row (Day 00)
        html += `<div class="sheet-cell initial-row-cell">0</div>`;
        html += `<div class="sheet-cell" style="display:none;"><input type="time" class="row-input" data-date="${monthKey}-00" data-field="hora" value="${initialLog.hora || ''}" ></div>`;
        equipment.forEach(e => {
            html += `<div class="sheet-cell"><input type="number" step="0.1" class="row-input" data-date="${monthKey}-00" data-field="h2_${e.id}_l" value="${initialLog['h2_' + e.id + '_l'] || ''}" ></div>`;
            html += `<div class="sheet-cell"><input type="number" step="0.1" class="row-input" data-date="${monthKey}-00" data-field="h2_${e.id}_h" value="${initialLog['h2_' + e.id + '_h'] || ''}" ></div>`;
        });

        const daysInMonth = new Date(this.currentYear, this.currentMonth + 1, 0).getDate();
        for (let d = 1; d <= 31; d++) {
            const dateStr = `${this.currentYear}-${String(this.currentMonth + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
            const log = logs.find(l => l.fecha === dateStr) || {};
            const isInvalidDay = d > daysInMonth;

            html += `<div class="sheet-cell date-cell ${isInvalidDay ? 'disabled-row' : ''}">${d}</div>`;
            html += `<div class="sheet-cell" style="display:none;"><input type="time" class="row-input" data-date="${dateStr}" data-field="hora" value="${log.hora || ''}" ${isInvalidDay ? 'disabled' : ''}></div>`;
            equipment.forEach(e => {
                html += `<div class="sheet-cell"><input type="number" step="0.1" class="row-input" data-date="${dateStr}" data-field="h2_${e.id}_l" value="${log['h2_' + e.id + '_l'] || ''}" ${isInvalidDay ? 'disabled' : ''}></div>`;
                html += `<div class="sheet-cell"><input type="number" step="0.1" class="row-input" data-date="${dateStr}" data-field="h2_${e.id}_h" value="${log['h2_' + e.id + '_h'] || ''}" ${isInvalidDay ? 'disabled' : ''}></div>`;
            });
        }

        html += `<div class="sheet-cell footer-cell span-col-2">TOTAL</div>`;
        equipment.forEach(e => {
            html += `<div class="sheet-cell footer-cell" id="total-h2-${e.id}-l">-</div><div class="sheet-cell footer-cell" id="total-h2-${e.id}-h">-</div>`;
        });
        html += `<div class="sheet-cell footer-cell span-col-2">MEDIA</div>`;
        equipment.forEach(e => {
            html += `<div class="sheet-cell footer-cell" id="media-h2-${e.id}-l">-</div><div class="sheet-cell footer-cell" id="media-h2-${e.id}-h">-</div>`;
        });
        return html + `</div>`;
    },

    renderCorredoiraHoja3Content() {
        const logs = JSON.parse(localStorage.getItem(`logs_EDAR_CORREDOIRA`) || '[]');
        const equipment = [
            { id: 'agita', label: 'AGITADOR A' }, { id: 'agitb', label: 'AGITADOR B' },
            { id: 'recba', label: 'REC. BIOLÓGICO A' }, { id: 'recbb', label: 'REC. BIOLÓGICO B' },
            { id: 'flotants', label: 'BOMBA FLOTANTES' }, { id: 'recfa', label: 'REC. FANGOS A' },
            { id: 'recfb', label: 'REC. FANGOS B' }, { id: 'sobrenad', label: 'SOBRENADANTES' }
        ];

        let html = `<div class="sheet-table corredoira-horas">
            <div class="sheet-cell sheet-header-cell span-row-2">DÍA</div>
            <div class="sheet-cell sheet-header-cell span-row-2" style="display:none;">HORA</div>`;
        equipment.forEach(e => { html += `<div class="sheet-cell sheet-header-cell span-col-2">${e.label}</div> `; });
        html += equipment.map(() => `<div class="sheet-cell sheet-header-cell">LECTURA</div><div class="sheet-cell sheet-header-cell">H.</div>`).join('');

        const monthKey = `${this.currentYear}-${String(this.currentMonth + 1).padStart(2, '0')}`;
        const initialLog = logs.find(l => l.fecha === `${monthKey}-00`) || {};

        // Initial Row (Day 00)
        html += `<div class="sheet-cell initial-row-cell">0</div>`;
        html += `<div class="sheet-cell" style="display:none;"><input type="time" class="row-input" data-date="${monthKey}-00" data-field="hora" value="${initialLog.hora || ''}" ></div>`;
        equipment.forEach(e => {
            html += `<div class="sheet-cell"><input type="number" step="0.1" class="row-input" data-date="${monthKey}-00" data-field="h3_${e.id}_l" value="${initialLog['h3_' + e.id + '_l'] || ''}" ></div>`;
            html += `<div class="sheet-cell"><input type="number" step="0.1" class="row-input" data-date="${monthKey}-00" data-field="h3_${e.id}_h" value="${initialLog['h3_' + e.id + '_h'] || ''}" ></div>`;
        });

        const daysInMonth = new Date(this.currentYear, this.currentMonth + 1, 0).getDate();
        for (let d = 1; d <= 31; d++) {
            const dateStr = `${this.currentYear}-${String(this.currentMonth + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
            const log = logs.find(l => l.fecha === dateStr) || {};
            const isInvalidDay = d > daysInMonth;

            html += `<div class="sheet-cell date-cell ${isInvalidDay ? 'disabled-row' : ''}">${d}</div>`;
            html += `<div class="sheet-cell" style="display:none;"><input type="time" class="row-input" data-date="${dateStr}" data-field="hora" value="${log.hora || ''}" ${isInvalidDay ? 'disabled' : ''}></div>`;
            equipment.forEach(e => {
                html += `<div class="sheet-cell"><input type="number" step="0.1" class="row-input" data-date="${dateStr}" data-field="h3_${e.id}_l" value="${log['h3_' + e.id + '_l'] || ''}" ${isInvalidDay ? 'disabled' : ''}></div>`;
                html += `<div class="sheet-cell"><input type="number" step="0.1" class="row-input" data-date="${dateStr}" data-field="h3_${e.id}_h" value="${log['h3_' + e.id + '_h'] || ''}" ${isInvalidDay ? 'disabled' : ''}></div>`;
            });
        }

        html += `<div class="sheet-cell footer-cell span-col-2">TOTAL</div>`;
        equipment.forEach(e => {
            html += `<div class="sheet-cell footer-cell" id="total-h3-${e.id}-l">-</div><div class="sheet-cell footer-cell" id="total-h3-${e.id}-h">-</div>`;
        });
        html += `<div class="sheet-cell footer-cell span-col-2">MEDIA</div>`;
        equipment.forEach(e => {
            html += `<div class="sheet-cell footer-cell" id="media-h3-${e.id}-l">-</div><div class="sheet-cell footer-cell" id="media-h3-${e.id}-h">-</div>`;
        });
        return html + `</div>`;
    },

    setupCalculationListeners() {
        document.querySelectorAll('.double-field-row').forEach(row => {
            const type = row.getAttribute('data-type');
            const diffDisplay = row.querySelector('.diff-display') || document.createElement('div');
            if (!row.querySelector('.diff-display')) {
                diffDisplay.className = 'diff-display';
                row.appendChild(diffDisplay);
            }

            if (type === 'double') {
                const hrsInput = row.querySelector('.hrs-input');
                const prevHrs = parseFloat(hrsInput.getAttribute('data-prev')) || 0;
                hrsInput.addEventListener('input', () => {
                    const diff = (parseFloat(hrsInput.value) || 0) - prevHrs;
                    diffDisplay.textContent = diff.toFixed(1);
                    diffDisplay.style.color = diff < 0 ? '#ff3b30' : '#34c759';
                });
            } else {
                const valInput = row.querySelector('.val-input');
                const prevVal = parseFloat(valInput.getAttribute('data-prev')) || 0;
                valInput.addEventListener('input', () => {
                    const diff = (parseFloat(valInput.value) || 0) - prevVal;
                    diffDisplay.textContent = diff.toFixed(1);
                    diffDisplay.style.color = diff < 0 ? '#ff3b30' : '#34c759';
                });
            }
        });
    },

    recalculateEtapTotals() {
        if (this.currentStation !== 'ETAP') return;

        const daysInMonth = new Date(this.currentYear, this.currentMonth + 1, 0).getDate();
        const monthKey = `${this.currentYear}-${String(this.currentMonth + 1).padStart(2, '0')}`;

        for (let d = 1; d <= daysInMonth; d++) {
            const currentDayKey = `${monthKey}-${String(d).padStart(2, '0')}`;
            const prevDayKey = `${monthKey}-${String(d - 1).padStart(2, '0')}`;

            if (this.currentEtapTab === 'hoja1') {
                // Entrada Diff
                this.calculateFieldDiff(currentDayKey, prevDayKey, 'p1182_entrada', 'entrada_dif');
                // Salida Diff
                this.calculateFieldDiff(currentDayKey, prevDayKey, 'salida_caudal_val', 'salida_caudal_dif');
            } else if (this.currentEtapTab === 'hoja4') {
                // Horas Bombeo ETAP M1, M2, M3
                this.calculateFieldDiff(currentDayKey, prevDayKey, 'b_etap_m1', 'b_etap_dif1');
                this.calculateFieldDiff(currentDayKey, prevDayKey, 'b_etap_m2', 'b_etap_dif2');
                this.calculateFieldDiff(currentDayKey, prevDayKey, 'b_etap_m3', 'b_etap_dif3');
                // Lagazos II B1, B2
                this.calculateFieldDiff(currentDayKey, prevDayKey, 'b_lag2_b1', 'b_lag2_difb1');
                this.calculateFieldDiff(currentDayKey, prevDayKey, 'b_lag2_b2', 'b_lag2_difb2');
            } else if (this.currentEtapTab === 'hoja5') {
                // Horas Bombeo Lagazos I B1, B2, B3
                this.calculateFieldDiff(currentDayKey, prevDayKey, 'b_lag1_b1', 'b_lag1_difb1');
                this.calculateFieldDiff(currentDayKey, prevDayKey, 'b_lag1_b2', 'b_lag1_difb2');
                this.calculateFieldDiff(currentDayKey, prevDayKey, 'b_lag1_b3', 'b_lag1_difb3');
            }
        }
    },

    calculateFieldDiff(currentDayKey, prevDayKey, fieldName, diffName) {
        const curInput = document.querySelector(`.row-input[data-date="${currentDayKey}"][data-field="${fieldName}"]`);
        const prevInput = document.querySelector(`.row-input[data-date="${prevDayKey}"][data-field="${fieldName}"]`);
        const targetInput = document.querySelector(`.row-input[data-date="${currentDayKey}"][data-field="${diffName}"]`);

        if (curInput && prevInput && targetInput) {
            const cVal = parseFloat(curInput.value);
            const pVal = parseFloat(prevInput.value);
            if (!isNaN(cVal) && !isNaN(pVal)) {
                const diff = cVal - pVal;
                targetInput.value = diff.toFixed(1);
                this.saveSingleField(currentDayKey, diffName, diff);
            }
        }
    },


    renderEtapSheet(container) {
        if (!this.currentMonth) {
            const now = new Date();
            this.currentMonth = now.getMonth();
            this.currentYear = now.getFullYear();
        }
        if (!this.currentEtapTab) this.currentEtapTab = 'hoja1';

        const months = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"];

        let html = `
            <div class="premium-header">
                <div class="header-row">
                    <div class="header-brand">
                        <span class="brand-name">AQUADEZA</span>
                        <span class="brand-tagline">SERVICIO DE LALÍN</span>
                    </div>
                    <div class="header-station-box">
                        <h1 class="station-title">E.T.A.P.</h1>
                        <div class="sheet-info-badge">HOJA: ${this.currentEtapTab.replace('hoja', '')}</div>
                    </div>
                </div>
                <div class="header-controls-row">
                    <div class="control-pill">
                        <label>MES</label>
                        <select id="month-select">
                            ${months.map((m, i) => `<option value="${i}" ${i === this.currentMonth ? 'selected' : ''}>${m}</option>`).join('')}
                        </select>
                    </div>
                    <div class="control-pill">
                        <label>AÑO</label>
                        <select id="year-select">
                            ${this.generateYearOptions(this.currentYear)}
                        </select>
                    </div>
                </div>
            </div>
            
            <div class="sheet-tabs">
                <button class="tab-btn ${this.currentEtapTab === 'hoja1' ? 'active' : ''}" data-tab="hoja1">HOJA 1</button>
                <button class="tab-btn ${this.currentEtapTab === 'hoja2' ? 'active' : ''}" data-tab="hoja2">HOJA 2</button>
                <button class="tab-btn ${this.currentEtapTab === 'hoja3' ? 'active' : ''}" data-tab="hoja3">HOJA 3</button>
                <button class="tab-btn ${this.currentEtapTab === 'hoja4' ? 'active' : ''}" data-tab="hoja4">HOJA 4</button>
                <button class="tab-btn ${this.currentEtapTab === 'hoja5' ? 'active' : ''}" data-tab="hoja5">HOJA 5</button>
            </div>

            <div id="etap-sheet-content">
                ${this.renderEtapHojaContent()}
            </div>

            <div class="form-actions">
                <button type="submit" class="btn-primary">GUARDAR TODO EL MES</button>
            </div>
        `;
        container.innerHTML = html;
        this.setupMonthlyEventListeners();
        this.setupEtapTabListeners();
        // Recalculate will be called when inputs change
    },

    renderEtapHojaContent() {
        switch (this.currentEtapTab) {
            case 'hoja1': return this.renderEtapHoja1Content();
            case 'hoja2': return this.renderEtapHoja2Content();
            case 'hoja3': return this.renderEtapHoja3Content();
            case 'hoja4': return this.renderEtapHoja4Content();
            case 'hoja5': return this.renderEtapHoja5Content();
            default: return '<div style="padding:20px;">Próximamente...</div>';
        }
    },

    renderEtapHoja1Content() {
        const logs = JSON.parse(localStorage.getItem(`logs_ETAP`) || '[]');
        const monthKey = `${this.currentYear}-${String(this.currentMonth + 1).padStart(2, '0')}`;

        let html = `<div class="sheet-table etap-hoja1">
            <div class="sheet-cell sheet-header-cell span-row-2">DIA</div>
            <div class="sheet-cell sheet-header-cell span-row-2" style="display:none;">HORA</div>
            <div class="sheet-cell sheet-header-cell span-col-4">CAUDAL</div>
            <div class="sheet-cell sheet-header-cell span-row-2">P 1.18.1</div>
            <div class="sheet-cell sheet-header-cell span-row-2">P 1.18.2</div>
            <div class="sheet-cell sheet-header-cell span-row-2">P 1.18.3</div>
            <div class="sheet-cell sheet-header-cell span-col-4">ENERGIA ACTIVA</div>
            <div class="sheet-cell sheet-header-cell">ENTRADA</div>
            <div class="sheet-cell sheet-header-cell">DIF</div>
            <div class="sheet-cell sheet-header-cell">SALIDA</div>
            <div class="sheet-cell sheet-header-cell">DIF</div>
            <div class="sheet-cell sheet-header-cell">P 1.18.4</div>
            <div class="sheet-cell sheet-header-cell">P 1.18.5</div>
            <div class="sheet-cell sheet-header-cell">P 1.18.6</div>
            <div class="sheet-cell sheet-header-cell">P 1.18.0</div>
        `;

        const daysInMonth = new Date(this.currentYear, this.currentMonth + 1, 0).getDate();
        for (let d = 0; d <= 31; d++) {
            const dateStr = `${monthKey}-${String(d).padStart(2, '0')}`;
            const log = logs.find(l => l.fecha === dateStr) || {};
            const isDisabled = d > daysInMonth && d !== 0;
            const rowClass = d === 0 ? 'initial-row-cell' : '';

            html += `<div class="sheet-cell ${rowClass} ${d > daysInMonth ? 'disabled-day' : ''}">${d}</div>`;
            html += `<div class="sheet-cell ${rowClass}"><input type="text" class="row-input" data-date="${dateStr}" data-field="hora" value="${log.hora || ''}" ${isDisabled ? 'disabled' : ''}></div>`;
            ['p1182_entrada', 'entrada_dif', 'salida_caudal_val', 'salida_caudal_dif', 'p1181', 'p1182', 'p1183', 'p1184', 'p1185', 'p1186', 'p1180'].forEach(f => {
                html += `<div class="sheet-cell ${rowClass}"><input type="number" step="0.1" class="row-input" data-date="${dateStr}" data-field="${f}" value="${log[f] || ''}" ${isDisabled ? 'disabled' : ''}></div>`;
            });
        }
        return html + `</div>`;
    },

    renderEtapHoja2Content() {
        const logs = JSON.parse(localStorage.getItem(`logs_ETAP`) || '[]');
        const monthKey = `${this.currentYear}-${String(this.currentMonth + 1).padStart(2, '0')}`;

        let html = `<div class="sheet-table etap-hoja2">
            <div class="sheet-cell sheet-header-cell span-row-2">DIA</div>
            <div class="sheet-cell sheet-header-cell span-row-2" style="display:none;">HORA</div>
            <div class="sheet-cell sheet-header-cell span-col-7">ENERGIA REACTIVA</div>
            <div class="sheet-cell sheet-header-cell span-col-7">MAXIMETRO</div>
            <div class="sheet-cell sheet-header-cell">P.1.58.1</div>
            <div class="sheet-cell sheet-header-cell">P.1.58.2</div>
            <div class="sheet-cell sheet-header-cell">P.1.58.3</div>
            <div class="sheet-cell sheet-header-cell">P.1.58.4</div>
            <div class="sheet-cell sheet-header-cell">P.1.58.5</div>
            <div class="sheet-cell sheet-header-cell">P.1.58.6</div>
            <div class="sheet-cell sheet-header-cell">P.1.58.0</div>
            <div class="sheet-cell sheet-header-cell">P.1.16.1</div>
            <div class="sheet-cell sheet-header-cell">P.1.16.2</div>
            <div class="sheet-cell sheet-header-cell">P.1.16.3</div>
            <div class="sheet-cell sheet-header-cell">P.1.16.4</div>
            <div class="sheet-cell sheet-header-cell">P.1.16.5</div>
            <div class="sheet-cell sheet-header-cell">P.1.16.6</div>
            <div class="sheet-cell sheet-header-cell">P.1.16.0</div>
        `;
        const daysInMonth = new Date(this.currentYear, this.currentMonth + 1, 0).getDate();
        for (let d = 0; d <= 31; d++) {
            const dateStr = `${monthKey}-${String(d).padStart(2, '0')}`;
            const log = logs.find(l => l.fecha === dateStr) || {};
            const isDisabled = d > daysInMonth && d !== 0;
            const rowClass = d === 0 ? 'initial-row-cell' : '';
            html += `<div class="sheet-cell ${rowClass} ${d > daysInMonth ? 'disabled-day' : ''}">${d}</div>`;
            html += `<div class="sheet-cell ${rowClass}"><input type="text" class="row-input" data-date="${dateStr}" data-field="hora" value="${log.hora || ''}" ${isDisabled ? 'disabled' : ''}></div>`;
            ['p1581', 'p1582', 'p1583', 'p1584', 'p1585', 'p1586', 'p1580', 'p1161', 'p1162', 'p1163', 'p1164', 'p1165', 'p1166', 'p1160'].forEach(f => {
                html += `<div class="sheet-cell ${rowClass}"><input type="number" step="0.1" class="row-input" data-date="${dateStr}" data-field="${f}" value="${log[f] || ''}" ${isDisabled ? 'disabled' : ''}></div>`;
            });
        }
        return html + `</div>`;
    },

    renderEtapHoja3Content() {
        const logs = JSON.parse(localStorage.getItem(`logs_ETAP`) || '[]');
        const monthKey = `${this.currentYear}-${String(this.currentMonth + 1).padStart(2, '0')}`;
        let html = `<div class="sheet-table etap-hoja3">
            <div class="sheet-cell sheet-header-cell span-row-2">DIA</div>
            <div class="sheet-cell sheet-header-cell span-row-2" style="display:none;">HORA</div>
            <div class="sheet-cell sheet-header-cell span-row-2"><div class="header-vertical">ALTURA<br>DPTO. ETAP</div></div>
            <div class="sheet-cell sheet-header-cell span-col-5">REGULACIÓN DOSIFICACIÓN REACTIVOS</div>
            <div class="sheet-cell sheet-header-cell"><div class="header-vertical">ph</div></div>
            <div class="sheet-cell sheet-header-cell"><div class="header-vertical">TURBIDEZ</div></div>
            <div class="sheet-cell sheet-header-cell"><div class="header-vertical">COLOR</div></div>
            <div class="sheet-cell sheet-header-cell"><div class="header-vertical">AMONIO</div></div>
            <div class="sheet-cell sheet-header-cell"><div class="header-vertical">Prep. react</div></div>
            <div class="sheet-cell sheet-header-cell span-col-4">STOCK REACTIVOS</div>
            <div class="sheet-cell sheet-header-cell">POSCLO.</div>
            <div class="sheet-cell sheet-header-cell">PRECLO.</div>
            <div class="sheet-cell sheet-header-cell">POLICLO.</div>
            <div class="sheet-cell sheet-header-cell">ALMIDON</div>
            <div class="sheet-cell sheet-header-cell">SOSA</div>
            <div class="sheet-cell sheet-header-cell">SALIDA</div>
            <div class="sheet-cell sheet-header-cell">SALIDA</div>
            <div class="sheet-cell sheet-header-cell">SALIDA</div>
            <div class="sheet-cell sheet-header-cell">SALIDA</div>
            <div class="sheet-cell sheet-header-cell">amonio</div>
            <div class="sheet-cell sheet-header-cell">HIPOCLO.</div>
            <div class="sheet-cell sheet-header-cell">POLICLO.</div>
            <div class="sheet-cell sheet-header-cell">ALMIDON</div>
            <div class="sheet-cell sheet-header-cell">SOSA</div>
        `;
        const daysInMonth = new Date(this.currentYear, this.currentMonth + 1, 0).getDate();
        for (let d = 0; d <= 31; d++) {
            const dateStr = `${monthKey}-${String(d).padStart(2, '0')}`;
            const log = logs.find(l => l.fecha === dateStr) || {};
            const isDisabled = d > daysInMonth && d !== 0;
            const rowClass = d === 0 ? 'initial-row-cell' : '';
            html += `<div class="sheet-cell ${rowClass} ${d > daysInMonth ? 'disabled-day' : ''}">${d}</div>`;
            html += `<div class="sheet-cell ${rowClass}"><input type="text" class="row-input" data-date="${dateStr}" data-field="hora" value="${log.hora || ''}" ${isDisabled ? 'disabled' : ''}></div>`;
            html += `<div class="sheet-cell ${rowClass}"><input type="number" step="0.01" class="row-input" data-date="${dateStr}" data-field="altura_etap" value="${log.altura_etap || ''}" ${isDisabled ? 'disabled' : ''}></div>`;
            ['dosif_pos', 'dosif_pre', 'dosif_poli', 'dosif_almi', 'dosif_sosa', 'ph_salida', 'turb_salida', 'color_salida', 'amonio_salida', 'prep_react_amonio', 'stock_hipo', 'stock_poli', 'stock_almi', 'stock_sosa'].forEach(f => {
                html += `<div class="sheet-cell ${rowClass}"><input type="number" step="0.001" class="row-input" data-date="${dateStr}" data-field="${f}" value="${log[f] || ''}" ${isDisabled ? 'disabled' : ''}></div>`;
            });
        }
        return html + `</div>`;
    },

    renderEtapHoja4Content() {
        const logs = JSON.parse(localStorage.getItem(`logs_ETAP`) || '[]');
        const monthKey = `${this.currentYear}-${String(this.currentMonth + 1).padStart(2, '0')}`;
        let html = `<div class="sheet-table etap-hoja4">
            <div class="sheet-cell sheet-header-cell span-row-2">DIA</div>
            <div class="sheet-cell sheet-header-cell span-row-2" style="display:none;">HORA</div>
            <div class="sheet-cell sheet-header-cell span-col-2">CL ETAP</div>
            <div class="sheet-cell sheet-header-cell span-col-2">DPTO LAGAZOS I</div>
            <div class="sheet-cell sheet-header-cell span-col-4">DPTO LAGAZOS II</div>
            <div class="sheet-cell sheet-header-cell span-col-6">HORAS BOMBEO ETAP</div>
            <div class="sheet-cell sheet-header-cell span-col-4">HORAS BOMBEO LAGAZOS II</div>
            <div class="sheet-cell sheet-header-cell">LIBRE</div>
            <div class="sheet-cell sheet-header-cell">COMB.</div>
            <div class="sheet-cell sheet-header-cell">ALTURA</div>
            <div class="sheet-cell sheet-header-cell">LIBRE</div>
            <div class="sheet-cell sheet-header-cell">ALTURA</div>
            <div class="sheet-cell sheet-header-cell">LIBRE</div>
            <div class="sheet-cell sheet-header-cell">Hipoc.</div>
            <div class="sheet-cell sheet-header-cell">Dilucion</div>
            <div class="sheet-cell sheet-header-cell">M1</div>
            <div class="sheet-cell sheet-header-cell">DIF.</div>
            <div class="sheet-cell sheet-header-cell">M2</div>
            <div class="sheet-cell sheet-header-cell">DIF.</div>
            <div class="sheet-cell sheet-header-cell">M3</div>
            <div class="sheet-cell sheet-header-cell">DIF.</div>
            <div class="sheet-cell sheet-header-cell">B1</div>
            <div class="sheet-cell sheet-header-cell">DIF.</div>
            <div class="sheet-cell sheet-header-cell">B2</div>
            <div class="sheet-cell sheet-header-cell">DIF.</div>
        `;
        const daysInMonth = new Date(this.currentYear, this.currentMonth + 1, 0).getDate();
        for (let d = 0; d <= 31; d++) {
            const dateStr = `${monthKey}-${String(d).padStart(2, '0')}`;
            const log = logs.find(l => l.fecha === dateStr) || {};
            const isDisabled = d > daysInMonth && d !== 0;
            const rowClass = d === 0 ? 'initial-row-cell' : '';
            html += `<div class="sheet-cell ${rowClass} ${d > daysInMonth ? 'disabled-day' : ''}">${d}</div>`;
            html += `<div class="sheet-cell ${rowClass}"><input type="text" class="row-input" data-date="${dateStr}" data-field="hora" value="${log.hora || ''}" ${isDisabled ? 'disabled' : ''}></div>`;
            ['cl_lib_etap', 'cl_comb_etap', 'lag1_alt', 'lag1_cl', 'lag2_alt', 'lag2_cl', 'lag2_hipo', 'lag2_dil', 'b_etap_m1', 'b_etap_dif1', 'b_etap_m2', 'b_etap_dif2', 'b_etap_m3', 'b_etap_dif3', 'b_lag2_b1', 'b_lag2_difb1', 'b_lag2_b2', 'b_lag2_difb2'].forEach(f => {
                html += `<div class="sheet-cell ${rowClass}"><input type="number" step="0.1" class="row-input" data-date="${dateStr}" data-field="${f}" value="${log[f] || ''}" ${isDisabled ? 'disabled' : ''}></div>`;
            });
        }
        return html + `</div>`;
    },

    renderEtapHoja5Content() {
        const logs = JSON.parse(localStorage.getItem(`logs_ETAP`) || '[]');
        const monthKey = `${this.currentYear}-${String(this.currentMonth + 1).padStart(2, '0')}`;
        let html = `<div class="sheet-table etap-hoja5">
            <div class="sheet-cell sheet-header-cell span-row-2">DIA</div>
            <div class="sheet-cell sheet-header-cell span-row-2" style="display:none;">HORA</div>
            <div class="sheet-cell sheet-header-cell span-row-2">LAGAZOS I<br>CONTADOR</div>
            <div class="sheet-cell sheet-header-cell span-col-6">HORAS BOMBEO LAGAZOS I</div>
            <div class="sheet-cell sheet-header-cell span-col-6">FILTROS LIMPIEZA</div>
            <div class="sheet-cell sheet-header-cell">B1</div>
            <div class="sheet-cell sheet-header-cell">DIF.</div>
            <div class="sheet-cell sheet-header-cell">B2</div>
            <div class="sheet-cell sheet-header-cell">DIF.</div>
            <div class="sheet-cell sheet-header-cell">B3</div>
            <div class="sheet-cell sheet-header-cell">DIF.</div>
            <div class="sheet-cell sheet-header-cell"><div class="header-vertical">analiz cl</div></div>
            <div class="sheet-cell sheet-header-cell"><div class="header-vertical">turb entr</div></div>
            <div class="sheet-cell sheet-header-cell"><div class="header-vertical">turb salid</div></div>
            <div class="sheet-cell sheet-header-cell"><div class="header-vertical">ph salid</div></div>
            <div class="sheet-cell sheet-header-cell"><div class="header-vertical">ph tratad</div></div>
            <div class="sheet-cell sheet-header-cell"><div class="header-vertical">amonio</div></div>
        `;
        const daysInMonth = new Date(this.currentYear, this.currentMonth + 1, 0).getDate();
        for (let d = 0; d <= 31; d++) {
            const dateStr = `${monthKey}-${String(d).padStart(2, '0')}`;
            const log = logs.find(l => l.fecha === dateStr) || {};
            const isDisabled = d > daysInMonth && d !== 0;
            const rowClass = d === 0 ? 'initial-row-cell' : '';
            html += `<div class="sheet-cell ${rowClass} ${d > daysInMonth ? 'disabled-day' : ''}">${d}</div>`;
            html += `<div class="sheet-cell ${rowClass}"><input type="text" class="row-input" data-date="${dateStr}" data-field="hora" value="${log.hora || ''}" ${isDisabled ? 'disabled' : ''}></div>`;
            html += `<div class="sheet-cell ${rowClass}"><input type="number" step="1" class="row-input" data-date="${dateStr}" data-field="lag1_contador" value="${log.lag1_contador || ''}" ${isDisabled ? 'disabled' : ''}></div>`;
            ['b_lag1_b1', 'b_lag1_difb1', 'b_lag1_b2', 'b_lag1_difb2', 'b_lag1_b3', 'b_lag1_difb3', 'f_anal_cl', 'f_turb_entr', 'f_turb_sal', 'f_ph_sal', 'f_ph_trat', 'f_amonio'].forEach(f => {
                html += `<div class="sheet-cell ${rowClass}"><input type="number" step="0.1" class="row-input" data-date="${dateStr}" data-field="${f}" value="${log[f] || ''}" ${isDisabled ? 'disabled' : ''}></div>`;
            });
        }
        return html + `</div>`;
    },

    setupEtapTabListeners() {
        document.querySelectorAll('.tab-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                this.currentEtapTab = e.target.getAttribute('data-tab');
                this.renderEtapSheet(document.getElementById('form-fields'));
            });
        });
    },


    handleFormSubmit() {
        if (['EDAR_BOTOS', 'EDAR_CORREDOIRA', 'ETAP', 'BOMBEO_BOTOS', 'CATASOS', 'VILATUXE'].includes(this.currentStation)) {
            const daysInMonth = new Date(this.currentYear, this.currentMonth + 1, 0).getDate();
            for (let d = 0; d <= daysInMonth; d++) {
                const dateStr = `${this.currentYear}-${String(this.currentMonth + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
                const rowInputs = document.querySelectorAll(`.row-input[data-date="${dateStr}"]`);

                const rowData = { fecha: dateStr };
                let hasData = false;
                rowInputs.forEach(input => {
                    const field = input.getAttribute('data-field');
                    const val = input.value;
                    if (val !== '') {
                        rowData[field] = input.type === 'number' ? parseFloat(val) : val;
                        hasData = true;
                    }
                });

                if (hasData) {
                    this.DataManager.saveEntry(this.currentStation, rowData);
                }
            }
            alert('Mes completo guardado correctamente');
        } else {
            const formData = {
                fecha: new Date().toISOString().split('T')[0],
                hora: new Date().toTimeString().substring(0, 5),
                observaciones: document.getElementById('form-obs')?.value || ''
            };

            document.querySelectorAll('.double-field-row').forEach(row => {
                const fieldName = row.getAttribute('data-field');
                const type = row.getAttribute('data-type');
                if (type === 'double') {
                    formData[fieldName] = {
                        lect: parseFloat(row.querySelector('.lect-input').value) || 0,
                        hrs: parseFloat(row.querySelector('.hrs-input').value) || 0
                    };
                } else {
                    formData[fieldName] = parseFloat(row.querySelector('.val-input').value) || 0;
                }
            });

            const precip = document.getElementById('field-precip');
            if (precip) formData.precipitacion = precip.value;

            const obs = document.getElementById('form-obs');
            if (obs) formData.observaciones = obs.value;

            this.DataManager.saveEntry(this.currentStation, formData);
            alert('Datos guardados correctamente');
        }
        this.showView('dashboard');
    },

    renderVilatuxeSheet(container) {
        if (!this.currentMonth) {
            const now = new Date();
            this.currentMonth = now.getMonth();
            this.currentYear = now.getFullYear();
        }
        const months = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"];

        let html = `
            <div class="premium-header">
                <div class="header-row">
                    <div class="header-brand">
                        <span class="brand-name">AQUADEZA</span>
                        <span class="brand-tagline">SERVICIO DE LALÍN</span>
                    </div>
                    <div class="header-station-box">
                        <h1 class="station-title">BOMBEO DE VILATUXE</h1>
                        <div class="sheet-info-badge">${this.currentVilatuxeTab.toUpperCase()}</div>
                    </div>
                </div>
                <div class="header-controls-row">
                    <div class="control-pill">
                        <label>MES</label>
                        <select id="month-select">
                            ${months.map((m, i) => `<option value="${i}" ${i === this.currentMonth ? 'selected' : ''}>${m}</option>`).join('')}
                        </select>
                    </div>
                    <div class="control-pill">
                        <label>AÑO</label>
                        <select id="year-select">
                            ${this.generateYearOptions(this.currentYear)}
                        </select>
                    </div>
                </div>
            </div>

            <div class="sheet-tabs">
                <button class="tab-btn ${this.currentVilatuxeTab === 'hoja1' ? 'active' : ''}" data-tab="hoja1">HOJA 1 (DIARIO)</button>
                <button class="tab-btn ${this.currentVilatuxeTab === 'hoja2' ? 'active' : ''}" data-tab="hoja2">HOJA 2 (ENERGÍA)</button>
            </div>

            <div id="vilatuxe-sheet-content">
                ${this.currentVilatuxeTab === 'hoja1' ? this.renderVilatuxeHoja1Content() : ''}
                ${this.currentVilatuxeTab === 'hoja2' ? this.renderVilatuxeHoja2Content() : ''}
            </div>
            <div style="margin-top:20px; text-align:right;">
                <button type="submit" class="btn-primary" style="width:auto; padding:10px 30px;">GUARDAR TODO EL MES</button>
            </div>
        `;
        container.innerHTML = html;
        this.setupMonthlyEventListeners();
        this.setupVilatuxeTabListeners();
        this.recalculateTotals();
    },

    setupVilatuxeTabListeners() {
        document.querySelectorAll('.tab-btn').forEach(btn => {
            btn.onclick = () => {
                this.currentVilatuxeTab = btn.dataset.tab;
                this.renderVilatuxeSheet(document.getElementById('form-fields'));
            };
        });
    },

    renderVilatuxeHoja1Content() {
        const logs = JSON.parse(localStorage.getItem(`logs_VILATUXE`) || '[]');
        const monthKey = `${this.currentYear}-${String(this.currentMonth + 1).padStart(2, '0')}`;
        let html = `<div class="sheet-table bombeo-catasos-hoja1">
            <div class="sheet-cell sheet-header-cell span-row-2">DIA</div>
            <div class="sheet-cell sheet-header-cell span-row-2" style="display:none;">HORA</div>
            <div class="sheet-cell sheet-header-cell span-col-6">HORAS BOMBEO</div>
            <div class="sheet-cell sheet-header-cell span-col-7">ENERGIA ACTIVA</div>
            <div class="sheet-cell sheet-header-cell">BOMBA 1</div><div class="sheet-cell sheet-header-cell">DIF</div>
            <div class="sheet-cell sheet-header-cell">BOMBA 2</div><div class="sheet-cell sheet-header-cell">DIF</div>
            <div class="sheet-cell sheet-header-cell">BOMBA 3</div><div class="sheet-cell sheet-header-cell">DIF</div>
            <div class="sheet-cell sheet-header-cell">P 1.18.1</div><div class="sheet-cell sheet-header-cell">P 1.18.2</div>
            <div class="sheet-cell sheet-header-cell">P 1.18.3</div><div class="sheet-cell sheet-header-cell">P 1.18.4</div>
            <div class="sheet-cell sheet-header-cell">P 1.18.5</div><div class="sheet-cell sheet-header-cell">P 1.18.6</div>
            <div class="sheet-cell sheet-header-cell">TOTAL</div>
        `;
        const daysInMonth = new Date(this.currentYear, this.currentMonth + 1, 0).getDate();
        for (let d = 0; d <= 31; d++) {
            const dateStr = `${monthKey}-${String(d).padStart(2, '0')}`;
            const log = logs.find(l => l.fecha === dateStr) || {};
            const isDisabled = d > daysInMonth;
            html += `
                <div class="sheet-cell ${d === 0 ? 'initial-row-cell' : ''}">${d}</div>
                <div class="sheet-cell" style="display:none;"><input type="time" class="row-input" data-date="${dateStr}" data-field="hora" value="${log.hora || ''}" ${isDisabled ? 'disabled' : ''}></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="b1" value="${log.b1 || ''}"></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="b1_dif" value="${log.b1_dif || ''}" disabled></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="b2" value="${log.b2 || ''}"></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="b2_dif" value="${log.b2_dif || ''}" disabled></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="b3" value="${log.b3 || ''}"></div>
                <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="b3_dif" value="${log.b3_dif || ''}" disabled></div>
                ${['p1181', 'p1182', 'p1183', 'p1184', 'p1185', 'p1186', 'p1180'].map(f => `
                    <div class="sheet-cell"><input type="number" class="row-input" data-date="${dateStr}" data-field="${f}" value="${log[f] || ''}"></div>
                `).join('')}
            `;
        }
        return html + `</div>`;
    },

    renderVilatuxeHoja2Content() {
        const logs = JSON.parse(localStorage.getItem(`logs_VILATUXE`) || '[]');
        return this.renderGenericEnergiaContent(logs, 'VILATUXE');
    },

    renderEdarCorredoiraEnergiaContent() {
        const logs = JSON.parse(localStorage.getItem(`logs_EDAR_CORREDOIRA`) || '[]');
        return this.renderGenericEnergiaContent(logs, 'EDAR_CORREDOIRA');
    },

    renderEdarCorredoiraEnergiaDifs() {
        // Reuse logic from Botos
        this.recalculateEnergiaDifs('EDAR_CORREDOIRA');
    },

    recalculateEnergiaDifs(station) {
        const monthKey = `${this.currentYear}-${String(this.currentMonth + 1).padStart(2, '0')}`;
        for (let d = 1; d <= 31; d++) {
            const cur = `${monthKey}-${String(d).padStart(2, '0')}`;
            const pre = `${monthKey}-${String(d - 1).padStart(2, '0')}`;
            ['e_a_total', 'e_r_total'].forEach(field => {
                const curIn = document.querySelector(`.row-input[data-date="${cur}"][data-field="${field}"]`);
                const preIn = document.querySelector(`.row-input[data-date="${pre}"][data-field="${field}"]`);
                const difIn = document.querySelector(`.row-input[data-date="${cur}"][data-field="${field === 'e_a_total' ? 'e_a_dif' : 'e_r_dif'}"]`);
                if (curIn && preIn && difIn) {
                    const v1 = parseFloat(curIn.value), v2 = parseFloat(preIn.value);
                    if (!isNaN(v1) && !isNaN(v2)) difIn.value = (v1 - v2).toFixed(0);
                }
            });
        }
    },

    async exportAllData(format) {
        const stations = ['ETAP', 'CATASOS', 'VILATUXE', 'BOMBEO_BOTOS', 'EDAR_BOTOS', 'EDAR_CORREDOIRA'];
        const allData = {};
        const monthKey = `${this.currentYear}-${String(this.currentMonth + 1).padStart(2, '0')}`;

        stations.forEach(s => {
            const logs = JSON.parse(localStorage.getItem(`logs_${s}`) || '[]');
            // Filter by selected year and month
            allData[s] = logs.filter(l => l.fecha && l.fecha.startsWith(monthKey));
        });

        if (format === 'excel') {
            this.generateExcel(allData);
        } else {
            this.generatePDF(allData);
        }
    },

    getSheetDefinitions(station) {
        if (station === 'EDAR_CORREDOIRA') {
            return [
                { name: 'DIARIO', headers: ['Día', 'Hora', 'Aire', 'Agua', 'Precip.', 'Lectura', 'm³/día', 'Oxíg.', 'V.Fango', 'Desb.', 'Arenas', 'Fangos', 'Punta', 'Valle', 'Llano', 'Reactiva', 'Obs.'], keys: ['fecha', 'hora', 'taire', 'tagua', 'precipitacion', 'caudal_lect', 'caudal_m3d', 'oxigeno', 'vol_fango', 'res_desbaste', 'res_arenas', 'res_fangos', 'cons_punta', 'cons_valle', 'cons_llano', 'reactiva', 'observaciones'] },
                { name: 'ENERGIA', headers: ['Día', 'P1 Act', 'P2 Act', 'P3 Act', 'P4 Act', 'P5 Act', 'P6 Act', 'Total Act', 'Dif Act', 'P1 Reac', 'P2 Reac', 'P3 Reac', 'P4 Reac', 'P5 Reac', 'P6 Reac', 'Total Reac', 'Dif Reac', 'M1', 'M2', 'M3', 'M Total'], keys: ['fecha', 'e_a_p1', 'e_a_p2', 'e_a_p3', 'e_a_p4', 'e_a_p5', 'e_a_p6', 'e_a_total', 'e_a_dif', 'e_r_p1', 'e_r_p2', 'e_r_p3', 'e_r_p4', 'e_r_p5', 'e_r_p6', 'e_r_total', 'e_r_dif', 'm_p1', 'm_p2', 'm_p3', 'm_total'] },
                { name: 'HORAS 1', headers: ['Día', 'Hora', 'Cinta L', 'Cinta H', 'Desar L', 'Desar H', 'B.Arenas L', 'B.Arenas H', 'Aerad L', 'Aerad H', 'Clasif L', 'Clasif H', 'Vaciad L', 'Vaciad H', 'Espes L', 'Espes H', 'Filtro L', 'Filtro H'], keys: ['fecha', 'hora', 'h1_cinta_l', 'h1_cinta_h', 'h1_desar_l', 'h1_desar_h', 'h1_barenas_l', 'h1_barenas_h', 'h1_aerad_l', 'h1_aerad_h', 'h1_careas_l', 'h1_careas_h', 'h1_vaciados_l', 'h1_vaciados_h', 'h1_espesador_l', 'h1_espesador_h', 'h1_filtro_l', 'h1_filtro_h'] },
                { name: 'HORAS 2', headers: ['Día', 'Hora', 'Turb A L', 'Turb A H', 'Turb B L', 'Turb B H', 'Turb C L', 'Turb C H', 'Turb D L', 'Turb D H', 'Turb E L', 'Turb E H', 'Turb F L', 'Turb F H', 'Turb G L', 'Turb G H'], keys: ['fecha', 'hora', 'h2_turba_l', 'h2_turba_h', 'h2_turbb_l', 'h2_turbb_h', 'h2_turbc_l', 'h2_turbc_h', 'h2_turbd_l', 'h2_turbd_h', 'h2_turbe_l', 'h2_turbe_h', 'h2_turbf_l', 'h2_turbf_h', 'h2_turbg_l', 'h2_turbg_h'] },
                { name: 'HORAS 3', headers: ['Día', 'Hora', 'Agit A L', 'Agit A H', 'Agit B L', 'Agit B H', 'Rec Biol A L', 'Rec Biol A H', 'Rec Biol B L', 'Rec Biol B H', 'Flotant L', 'Flotant H', 'Rec Fan A L', 'Rec Fan A H', 'Rec Fan B L', 'Rec Fan B H', 'Sobrenad L', 'Sobrenad H'], keys: ['fecha', 'hora', 'h3_agita_l', 'h3_agita_h', 'h3_agitb_l', 'h3_agitb_h', 'h3_recba_l', 'h3_recba_h', 'h3_recbb_l', 'h3_recbb_h', 'h3_flotants_l', 'h3_flotants_h', 'h3_recfa_l', 'h3_recfa_h', 'h3_recfb_l', 'h3_recfb_h', 'h3_sobrenad_l', 'h3_sobrenad_h'] }
            ];
        } else if (station === 'EDAR_BOTOS') {
            return [
                { name: 'DIARIO', headers: ['Día', 'Hora', 'Aire', 'Agua', 'Precip.', 'Lectura', 'm³/día', 'Desb.', 'Fangos', 'PAX', 'Poli.', 'Punta', 'Valle', 'Llano', 'Reactiva', 'Obs.'], keys: ['fecha', 'hora', 'taire', 'tagua', 'precipitacion', 'caudal_lect', 'caudal_m3d', 'desbaste', 'fangos_filtro', 'pax', 'polielectrolito', 'punta', 'valle', 'llano', 'reactiva', 'observaciones'] },
                { name: 'HORAS 1', headers: ['Día', 'Hora', 'B1 L', 'B1 H', 'B2 L', 'B2 H', 'Achique L', 'Achique H', 'Masko L', 'Masko H', 'Agit L', 'Agit H', 'Parrilla L', 'Parrilla H', 'Barred L', 'Barred H', 'Tornillo L', 'Tornillo H', 'Fango L', 'Fango H'], keys: ['fecha', 'hora', 'h2_b1_l', 'h2_b1_h', 'h2_b2_l', 'h2_b2_h', 'h2_achique_l', 'h2_achique_h', 'h2_masko_l', 'h2_masko_h', 'h2_agitador_l', 'h2_agitador_h', 'h2_parrilla_l', 'h2_parrilla_h', 'h2_barredera_l', 'h2_barredera_h', 'h2_tornillo_l', 'h2_tornillo_h', 'h2_fango_l', 'h2_fango_h'] },
                { name: 'HORAS 2', headers: ['Día', 'Hora', 'DAC1 L', 'DAC1 H', 'DAC2 L', 'DAC2 H', 'Comp L', 'Comp H', 'Floc1 L', 'Floc1 H', 'Floc2 L', 'Floc2 H', 'Coag1 L', 'Coag1 H', 'Coag2 L', 'Coag2 H', 'Sopl1 L', 'Sopl1 H', 'Sopl2 L', 'Sopl2 H'], keys: ['fecha', 'hora', 'h3_dac1_l', 'h3_dac1_h', 'h3_dac2_l', 'h3_dac2_h', 'h3_comp_l', 'h3_comp_h', 'h3_floc1_l', 'h3_floc1_h', 'h3_floc2_l', 'h3_floc2_h', 'h3_coag1_l', 'h3_coag1_h', 'h3_coag2_l', 'h3_coag2_h', 'h3_sopl1_l', 'h3_sopl1_h', 'h3_sopl2_l', 'h3_sopl2_h'] },
                { name: 'HORAS 3', headers: ['Día', 'Hora', 'Mono L', 'Mono H', 'Poli L', 'Poli H', 'Agit L', 'Agit H', 'Banda L', 'Banda H', 'V1 L', 'V1 H', 'V2 L', 'V2 H', 'V3 L', 'V3 H', 'V4 L', 'V4 H', 'V5 L', 'V5 H'], keys: ['fecha', 'hora', 'h4_mono_l', 'h4_mono_h', 'h4_poli_l', 'h4_poli_h', 'h4_agit_l', 'h4_agit_h', 'h4_banda_l', 'h4_banda_h', 'h4_v1_l', 'h4_v1_h', 'h4_v2_l', 'h4_v2_h', 'h4_v3_l', 'h4_v3_h', 'h4_v4_l', 'h4_v4_h', 'h4_v5_l', 'h4_v5_h'] },
                { name: 'ENERGIA', headers: ['Día', 'P1 Act', 'P2 Act', 'P3 Act', 'P4 Act', 'P5 Act', 'P6 Act', 'Total Act', 'Dif Act', 'P1 Reac', 'P2 Reac', 'P3 Reac', 'P4 Reac', 'P5 Reac', 'P6 Reac', 'Total Reac', 'Dif Reac', 'M1', 'M2', 'M3', 'M Total'], keys: ['fecha', 'e_a_p1', 'e_a_p2', 'e_a_p3', 'e_a_p4', 'e_a_p5', 'e_a_p6', 'e_a_total', 'e_a_dif', 'e_r_p1', 'e_r_p2', 'e_r_p3', 'e_r_p4', 'e_r_p5', 'e_r_p6', 'e_r_total', 'e_r_dif', 'm_p1', 'm_p2', 'm_p3', 'm_total'] }
            ];
        } else if (station === 'ETAP') {
            return [
                { name: 'HOJA 1', headers: ['Día', 'Hora', 'Entrada', 'Dif', 'Salida', 'Dif', 'P1.18.1', 'P1.18.2', 'P1.18.3', 'P1.18.4', 'P1.18.5', 'P1.18.6', 'P1.18.0'], keys: ['fecha', 'hora', 'p1182_entrada', 'entrada_dif', 'salida_caudal_val', 'salida_caudal_dif', 'p1181', 'p1182', 'p1183', 'p1184', 'p1185', 'p1186', 'p1180'] },
                { name: 'HOJA 2', headers: ['Día', 'Hora', 'P1.58.1', 'P1.58.2', 'P1.58.3', 'P1.58.4', 'P1.58.5', 'P1.58.6', 'P1.58.0', 'P1.16.1', 'P1.16.2', 'P1.16.3', 'P1.16.4', 'P1.16.5', 'P1.16.6', 'P1.16.0'], keys: ['fecha', 'hora', 'p1581', 'p1582', 'p1583', 'p1584', 'p1585', 'p1586', 'p1580', 'p1161', 'p1162', 'p1163', 'p1164', 'p1165', 'p1166', 'p1160'] },
                { name: 'HOJA 3', headers: ['Día', 'Hora', 'Altura', 'Pos', 'Pre', 'Poli', 'Almi', 'Sosa', 'pH', 'Turbid', 'Color', 'Ammon', 'Prep.', 'Stock H', 'Stock P', 'Stock A', 'Stock S'], keys: ['fecha', 'hora', 'altura_etap', 'dosif_pos', 'dosif_pre', 'dosif_poli', 'dosif_almi', 'dosif_sosa', 'ph_salida', 'turb_salida', 'color_salida', 'amonio_salida', 'prep_react_amonio', 'stock_hipo', 'stock_poli', 'stock_almi', 'stock_sosa'] },
                { name: 'HOJA 4', headers: ['Día', 'Hora', 'CL Lib', 'CL Comb', 'L1 Alt', 'L1 CL', 'L2 Alt', 'L2 CL', 'L2 Hipo', 'L2 Dil', 'M1', 'Dif', 'M2', 'Dif', 'M3', 'Dif', 'B1', 'Dif', 'B2', 'Dif'], keys: ['fecha', 'hora', 'cl_lib_etap', 'cl_comb_etap', 'lag1_alt', 'lag1_cl', 'lag2_alt', 'lag2_cl', 'lag2_hipo', 'lag2_dil', 'b_etap_m1', 'b_etap_dif1', 'b_etap_m2', 'b_etap_dif2', 'b_etap_m3', 'b_etap_dif3', 'b_lag2_b1', 'b_lag2_difb1', 'b_lag2_b2', 'b_lag2_difb2'] },
                { name: 'HOJA 5', headers: ['Día', 'Hora', 'L1 Cont', 'B1', 'Dif', 'B2', 'Dif', 'B3', 'Dif', 'Anal Cl', 'Turb E', 'Turb S', 'pH S', 'pH T', 'Ammon'], keys: ['fecha', 'hora', 'lag1_contador', 'b_lag1_b1', 'b_lag1_difb1', 'b_lag1_b2', 'b_lag1_difb2', 'b_lag1_b3', 'b_lag1_difb3', 'f_anal_cl', 'f_turb_entr', 'f_turb_sal', 'f_ph_sal', 'f_ph_trat', 'f_amonio'] }
            ];
        } else if (station === 'BOMBEO_BOTOS' || station === 'CATASOS' || station === 'VILATUXE') {
            return [
                { name: 'HOJA 1', headers: ['Día', 'Hora', 'B1', 'B1 Dif', 'B2', 'B2 Dif', 'B3', 'B3 Dif', 'P1.18.1', 'P1.18.2', 'P1.18.3', 'P1.18.4', 'P1.18.5', 'P1.18.6', 'P1.18.0'], keys: ['fecha', 'hora', 'b1', 'b1_dif', 'b2', 'b2_dif', 'b3', 'b3_dif', 'p1181', 'p1182', 'p1183', 'p1184', 'p1185', 'p1186', 'p1180'] },
                { name: 'HOJA 2', headers: ['Día', 'Hora', 'P1.58.1', 'P1.58.2', 'P1.58.3', 'P1.58.4', 'P1.58.5', 'P1.58.6', 'P1.58.0', 'P1.16.1', 'P1.16.2', 'P1.16.3', 'P1.16.4', 'P1.16.5', 'P1.16.6', 'P1.16.0'], keys: ['fecha', 'hora', 'p1581', 'p1582', 'p1583', 'p1584', 'p1585', 'p1586', 'p1580', 'p1161', 'p1162', 'p1163', 'p1164', 'p1165', 'p1166', 'p1160'] }
            ];
        }
        return [];
    },

    generateExcel(allData) {
        try {
            const wb = XLSX.utils.book_new();
            Object.keys(allData).forEach(station => {
                const logs = allData[station];
                const sheetDefs = this.getSheetDefinitions(station);

                sheetDefs.forEach(def => {
                    const aoaData = [def.headers]; // Add headers first

                    // Sort logs by date correctly
                    const sortedLogs = logs.sort((a, b) => (a.fecha || '').localeCompare(b.fecha || ''));

                    sortedLogs.forEach(row => {
                        const rowArray = def.keys.map(key => {
                            let val = row[key];
                            if (key === 'fecha' && val && val.includes('-')) {
                                const day = val.split('-').pop();
                                return day === '00' ? 'INICIAL' : parseInt(day);
                            }
                            if (val === undefined || val === null || val === '') return '-';
                            return typeof val === 'number' ? (Number.isInteger(val) ? val : parseFloat(val.toFixed(2))) : val;
                        });
                        aoaData.push(rowArray);
                    });

                    const ws = XLSX.utils.aoa_to_sheet(aoaData);
                    const safeSheetName = `${station.substring(0, 10)}_${def.name.substring(0, 15)}`.substring(0, 31);
                    XLSX.utils.book_append_sheet(wb, ws, safeSheetName);
                });
            });
            XLSX.writeFile(wb, `REGISTROS_AQUADEZA_${this.currentYear}_${String(this.currentMonth + 1).padStart(2, '0')}.xlsx`);
        } catch (error) {
            console.error("Excel Export Error:", error);
            alert("Error al exportar a Excel.");
        }
    },

    generatePDF(allData) {
        try {
            const { jsPDF } = window.jspdf;
            const doc = new jsPDF('l', 'mm', 'a4');
            let firstPage = true;

            Object.keys(allData).forEach(station => {
                const logs = allData[station];
                const sheetDefs = this.getSheetDefinitions(station);

                sheetDefs.forEach(sheet => {
                    if (!firstPage) doc.addPage();
                    firstPage = false;

                    doc.setFontSize(14);
                    doc.setTextColor(0, 56, 101);
                    doc.text(`${station.replace(/_/g, ' ')} - ${sheet.name} (${this.currentYear}-${String(this.currentMonth + 1).padStart(2, '0')})`, 14, 15);

                    doc.setFontSize(8);
                    doc.setTextColor(100);
                    doc.text(`Exportado: ${new Date().toLocaleDateString()} - Registro Diario`, 14, 21);

                    const body = logs.sort((a, b) => (a.fecha || '').localeCompare(b.fecha || '')).map(row => {
                        return sheet.keys.map(key => {
                            let val = row[key];
                            if (key === 'fecha' && val && val.includes('-')) {
                                const day = val.split('-').pop();
                                return day === '00' ? 'INIC.' : day;
                            }
                            if (val === undefined || val === null || val === '') return '-';
                            return typeof val === 'number' ? (Number.isInteger(val) ? val : val.toFixed(1)) : val;
                        });
                    });

                    doc.autoTable({
                        head: [sheet.headers],
                        body: body,
                        startY: 25,
                        theme: 'grid',
                        styles: { fontSize: 4.5, cellPadding: 0.5, halign: 'center' },
                        headStyles: { fillColor: [0, 56, 101], textColor: 255, fontSize: 4.5, fontStyle: 'bold' },
                        alternateRowStyles: { fillColor: [248, 249, 250] },
                        margin: { left: 8, right: 8 }
                    });
                });
            });

            doc.save(`REGISTROS_AQUADEZA_${this.currentYear}_${String(this.currentMonth + 1).padStart(2, '0')}.pdf`);
        } catch (error) {
            console.error("PDF Export Error:", error);
            alert("Error al exportar a PDF.");
        }
    }
};


console.log('App version: 3.6 - Year Selector Fixed');
App.init();
