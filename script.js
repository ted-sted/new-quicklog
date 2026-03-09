        // Placeholder for Application Logic (app.js)
        const app = {
            db: null,
            tokenClient: null,
            tokenExpiry: null,
            refreshTimer: null,
            timeToggled: false,
            clientId: '168669644981-at642rkfpbj1v4mbjtj1rd4osqn9uea9.apps.googleusercontent.com',
            currentEditId: null,
            pendingAuthAction: null,

            init: async function () {
                this.bindEvents();
                try {
                    await this.openDB();
                    await this.renderList();
                    this.updateSyncBadge();
                    document.getElementById('inputText').focus();

                    // Microphone permission cache (prevent startup beep)
                    navigator.mediaDevices.getUserMedia({ audio: true })
                        .then(stream => stream.getTracks().forEach(t => t.stop()))
                        .catch(() => { }); // ignore error

                    // Init Google Identity Services
                    this.initGIS();
                } catch (e) {
                    this.showError('繝・・繧ｿ繝ｼ繝吶・繧ｹ繧ｨ繝ｩ繝ｼ: ' + e.message);
                }
            },

            openDB: function () {
                return new Promise((resolve, reject) => {
                    // 繝舌・繧ｸ繝ｧ繝ｳ繧・縺ｫ荳翫￡縺ｦ縲∽ｻ･蜑阪・QuickLogDB縺後≠繧句ｴ蜷医〒繧ら｢ｺ螳溘↓繧｢繝・・繧ｰ繝ｬ繝ｼ繝会ｼ医う繝ｳ繝・ャ繧ｯ繧ｹ菴懈・・峨ｒ襍ｰ繧峨○繧・
                    const request = indexedDB.open('QuickLogDB', 2);

                    request.onupgradeneeded = (e) => {
                        const db = e.target.result;
                        let store;
                        if (!db.objectStoreNames.contains('entries')) {
                            store = db.createObjectStore('entries', { keyPath: 'id' });
                        } else {
                            store = e.target.transaction.objectStore('entries');
                        }

                        if (!store.indexNames.contains('timestamp')) {
                            store.createIndex('timestamp', 'timestamp', { unique: false });
                        }
                    };

                    request.onsuccess = (e) => {
                        this.db = e.target.result;
                        resolve(this.db);
                    };

                    request.onerror = (e) => {
                        reject(e.target.error);
                    };
                });
            },

            getAllEntries: function () {
                return new Promise((resolve, reject) => {
                    if (!this.db) return resolve([]);
                    try {
                        const transaction = this.db.transaction(['entries'], 'readonly');
                        const store = transaction.objectStore('entries');
                        const request = store.getAll(); // Bypass index to avoid NotFoundError on old DBs

                        request.onsuccess = () => {
                            let entries = request.result || [];
                            // Apply default values for old format compatibility
                            entries = entries.map(entry => ({
                                id: entry.id,
                                content: entry.content || '',
                                timestamp: entry.timestamp,
                                endTimestamp: entry.endTimestamp || null,
                                gcalId: entry.gcalId || null,
                                colorId: entry.colorId || null,
                                syncStatus: entry.syncStatus || 'pending'
                            }));
                            resolve(entries);
                        };
                        request.onerror = (e) => reject(e.target.error);
                    } catch (e) {
                        reject(e);
                    }
                });
            },

            saveEntry: function (entry) {
                return new Promise((resolve, reject) => {
                    const transaction = this.db.transaction(['entries'], 'readwrite');
                    const store = transaction.objectStore('entries');
                    const request = store.put(entry);

                    request.onsuccess = () => resolve();
                    request.onerror = (e) => reject(e.target.error);
                });
            },

            formatDateHeader: function (timestamp) {
                const date = new Date(timestamp);
                const today = new Date();
                const yesterday = new Date(today);
                yesterday.setDate(yesterday.getDate() - 1);

                const isSameDay = (d1, d2) =>
                    d1.getDate() === d2.getDate() &&
                    d1.getMonth() === d2.getMonth() &&
                    d1.getFullYear() === d2.getFullYear();

                if (isSameDay(date, today)) return '莉頑律';
                if (isSameDay(date, yesterday)) return '譏ｨ譌･';

                const days = ['譌･', '譛・, '轣ｫ', '豌ｴ', '譛ｨ', '驥・, '蝨・];
                return `${date.getMonth() + 1}/${date.getDate()}(${days[date.getDay()]})`;
            },

            formatTime: function (timestamp) {
                const date = new Date(timestamp);
                return `${date.getHours().toString().padStart(2, '0')}:${date.getMinutes().toString().padStart(2, '0')}`;
            },

            renderList: async function () {
                const listArea = document.getElementById('listArea');
                const emptyState = document.getElementById('emptyState');

                // Clear existing items except emptyState and importSection
                Array.from(listArea.children).forEach(child => {
                    if (child.id !== 'emptyState' && child.id !== 'importSection') child.remove();
                });

                const entries = await this.getAllEntries();

                // Filter out delete_pending
                const visibleEntries = entries.filter(e => e.syncStatus !== 'delete_pending');

                document.getElementById('headerTotalCount').textContent = visibleEntries.length + '莉ｶ';

                if (visibleEntries.length === 0) {
                    emptyState.style.display = 'flex';
                } else {
                    emptyState.style.display = 'none';
                }

                // 蛻晏屓繧､繝ｳ繝昴・繝医・繧ｿ繝ｳ縺ｮ陦ｨ遉ｺ蛻ｶ蠕｡・医☆縺ｧ縺ｫGCal逕ｱ譚･縺ｮ繝・・繧ｿ縺後≠繧後・髫縺呻ｼ・
                const hasImported = entries.some(e => e.gcalId && e.syncStatus === 'synced');
                document.getElementById('importSection').style.display = hasImported ? 'none' : 'block';

                const now = Date.now();
                const pastEntries = visibleEntries.filter(e => e.timestamp <= now).sort((a, b) => b.timestamp - a.timestamp);
                const futureEntries = visibleEntries.filter(e => e.timestamp > now).sort((a, b) => a.timestamp - b.timestamp); // 莉翫↓霑代＞莠亥ｮ壹°繧芽｡ｨ遉ｺ

                // Render Future Entries
                if (futureEntries.length > 0) {
                    const header = document.createElement('div');
                    header.className = 'date-header';
                    header.textContent = '竢ｰ 莠亥ｮ・;
                    listArea.appendChild(header);

                    futureEntries.forEach(entry => this.appendEntryCard(listArea, entry, true));
                }

                // Render Past Entries grouped by date
                let currentDateStr = '';
                pastEntries.forEach(entry => {
                    const dateStr = this.formatDateHeader(entry.timestamp);
                    if (dateStr !== currentDateStr) {
                        const header = document.createElement('div');
                        header.className = 'date-header';
                        header.textContent = dateStr;
                        listArea.appendChild(header);
                        currentDateStr = dateStr;
                    }
                    this.appendEntryCard(listArea, entry, false);
                });
            },

            appendEntryCard: function (container, entry, isFuture) {
                const wrapper = document.createElement('div');
                wrapper.className = 'entry-container';

                const deleteBg = document.createElement('div');
                deleteBg.className = 'delete-action-bg';
                deleteBg.textContent = '卵・・;

                const card = document.createElement('div');
                card.className = `entry-card ${isFuture ? 'future' : ''}`;
                card.dataset.id = entry.id;

                let timeStr = this.formatTime(entry.timestamp);
                if (entry.endTimestamp) {
                    timeStr += ` ・・${this.formatTime(entry.endTimestamp)}`;
                }

                const pendingHtml = entry.syncStatus === 'pending' ? '<span class="pending-mark">笳・/span>' : '';

                card.innerHTML = `
                    <div class="entry-time">${timeStr}</div>
                    <div class="entry-content">${this.escapeHTML(entry.content)}${pendingHtml}</div>
                    <button class="edit-btn" onclick="app.openEditModal('${entry.id}')">笨擾ｸ・/button>
                `;

                // Swipe logic
                let startX = 0;
                let currentX = 0;
                let isDragging = false;

                card.addEventListener('touchstart', (e) => {
                    startX = e.touches[0].clientX;
                    isDragging = true;
                    card.style.transition = 'none';
                }, { passive: true });

                card.addEventListener('touchmove', (e) => {
                    if (!isDragging) return;
                    currentX = e.touches[0].clientX;
                    const diffX = currentX - startX;
                    if (diffX < 0) { // dragging left
                        card.style.transform = `translateX(${Math.max(diffX, -100)}px)`;
                    }
                }, { passive: true });

                card.addEventListener('touchend', (e) => {
                    if (!isDragging) return;
                    isDragging = false;
                    card.style.transition = 'transform 0.3s ease';
                    const diffX = currentX - startX;
                    if (diffX < -70) {
                        card.style.transform = `translateX(-100%)`;
                        setTimeout(() => {
                            if (confirm('縺薙・險倬鹸繧貞炎髯､縺励∪縺吶°・・)) {
                                this.deleteEntry(entry.id);
                            } else {
                                card.style.transform = `translateX(0)`;
                            }
                        }, 300);
                    } else {
                        card.style.transform = `translateX(0)`;
                    }
                });

                // Tap to edit
                card.addEventListener('click', (e) => {
                    if (e.target.tagName !== 'BUTTON') {
                        this.openEditModal(entry.id);
                    }
                });

                wrapper.appendChild(deleteBg);
                wrapper.appendChild(card);
                container.appendChild(wrapper);
            },

            deleteEntry: async function (id) {
                try {
                    const entries = await this.getAllEntries();
                    const entry = entries.find(e => e.id === id);
                    if (!entry) return;

                    if (entry.syncStatus === 'synced' && entry.gcalId) {
                        entry.syncStatus = 'delete_pending';
                        await this.saveEntry(entry);
                    } else {
                        const transaction = this.db.transaction(['entries'], 'readwrite');
                        const store = transaction.objectStore('entries');
                        store.delete(id);
                        await new Promise((resolve, reject) => {
                            transaction.oncomplete = () => resolve();
                            transaction.onerror = (e) => reject(e.target.error);
                        });
                    }

                    await this.renderList();
                    this.updateSyncBadge();
                    this.showToast('蜑企勁縺励∪縺励◆');
                } catch (e) {
                    this.showError('蜑企勁縺ｫ螟ｱ謨励＠縺ｾ縺励◆');
                }
            },

            escapeHTML: function (str) {
                return str.replace(/[&<>'"]/g,
                    tag => ({
                        '&': '&amp;',
                        '<': '&lt;',
                        '>': '&gt;',
                        "'": '&#39;',
                        '"': '&quot;'
                    }[tag])
                );
            },

            updateSyncBadge: async function () {
                const entries = await this.getAllEntries();
                const pendingCount = entries.filter(e => e.syncStatus === 'pending' || e.syncStatus === 'delete_pending').length;

                const badge = document.getElementById('syncBadge');
                if (pendingCount > 0) {
                    badge.textContent = pendingCount;
                    badge.style.display = 'flex';
                } else {
                    badge.style.display = 'none';
                }
            },

            handleSend: async function () {
                const inputText = document.getElementById('inputText');
                const content = inputText.value.trim();

                if (!content) return;

                const isPlanMode = document.getElementById('btnModePlan').classList.contains('inactive') === false;

                let startTs, endTs = null;
                let colorId = isPlanMode ? '3' : null;

                if (isPlanMode || this.timeToggled) {
                    const startVal = document.getElementById('inputStartDt').value;
                    const endVal = document.getElementById('inputEndDt').value;

                    if (!startVal) {
                        this.showError('髢句ｧ区律譎ゅｒ蜈･蜉帙＠縺ｦ縺上□縺輔＞');
                        return;
                    }

                    startTs = new Date(startVal).getTime();
                    if (endVal) {
                        endTs = new Date(endVal).getTime();
                        if (endTs <= startTs) {
                            this.showError('邨ゆｺ・律譎ゅ・髢句ｧ区律譎ゅｈ繧雁ｾ後〒縺ゅｋ蠢・ｦ√′縺ゅｊ縺ｾ縺・);
                            return;
                        }
                    }
                } else {
                    startTs = Date.now();
                }

                const entry = {
                    id: `${startTs}_${Math.random().toString(36).substr(2, 9)}`,
                    content: content,
                    timestamp: startTs,
                    endTimestamp: endTs,
                    gcalId: null,
                    colorId: colorId,
                    syncStatus: 'pending'
                };

                try {
                    await this.saveEntry(entry);
                    inputText.value = '';
                    inputText.style.height = '40px';
                    document.getElementById('btnSend').disabled = true;

                    if (isPlanMode) {
                        // Switch back to normal mode
                        document.getElementById('btnModeNow').click();
                        this.showToast('莠亥ｮ壹ｒ霑ｽ蜉縺励∪縺励◆ 笨・);
                    } else {
                        this.timeToggled = false;
                        document.getElementById('datetimeInputs').classList.add('hidden');
                        document.getElementById('btnClock').style.backgroundColor = 'var(--bg-light)';
                        document.getElementById('btnClock').style.display = 'flex';
                        this.showToast('險倬鹸縺励∪縺励◆ 笨・);
                    }

                    await this.renderList();
                    this.updateSyncBadge();
                } catch (e) {
                    this.showError('菫晏ｭ倥↓螟ｱ謨励＠縺ｾ縺励◆');
                }
            },

            bindEvents: function () {
                // UI interaction bindings
                const inputText = document.getElementById('inputText');
                const btnSend = document.getElementById('btnSend');
                const btnModeNow = document.getElementById('btnModeNow');
                const btnModePlan = document.getElementById('btnModePlan');
                const datetimeInputs = document.getElementById('datetimeInputs');

                // Auto-resize textarea
                inputText.addEventListener('input', function () {
                    this.style.height = '40px';
                    this.style.height = Math.min(this.scrollHeight, 110) + 'px';
                    btnSend.disabled = this.value.trim().length === 0;
                });

                const setNowDefaults = () => {
                    const now = new Date();
                    const nowEnd = new Date(now.getTime() + 5 * 60 * 1000); // +5 mins
                    document.getElementById('inputStartDt').value = this.toLocalISOString(now).slice(0, 16);
                    document.getElementById('inputEndDt').value = this.toLocalISOString(nowEnd).slice(0, 16);
                };

                const btnClock = document.getElementById('btnClock');
                const updateClockVisibility = () => {
                    const isPlanMode = document.getElementById('btnModePlan').classList.contains('inactive') === false;
                    if (isPlanMode || this.timeToggled) {
                        datetimeInputs.classList.remove('hidden');
                    } else {
                        datetimeInputs.classList.add('hidden');
                    }
                    if (isPlanMode) {
                        btnClock.style.display = 'none';
                    } else {
                        btnClock.style.display = 'flex';
                        btnClock.style.backgroundColor = this.timeToggled ? 'var(--border-color)' : 'var(--bg-light)';
                    }
                };

                btnClock.addEventListener('click', () => {
                    this.timeToggled = !this.timeToggled;
                    if (this.timeToggled) {
                        setNowDefaults();
                    }
                    updateClockVisibility();
                });

                // Mode switching
                btnModeNow.addEventListener('click', () => {
                    btnModeNow.classList.remove('inactive');
                    btnModePlan.classList.add('inactive');
                    this.timeToggled = false;
                    setNowDefaults();
                    updateClockVisibility();
                });

                btnModePlan.addEventListener('click', () => {
                    btnModePlan.classList.remove('inactive');
                    btnModeNow.classList.add('inactive');

                    // Set default plan time: tomorrow same time
                    const tomorrow = new Date();
                    tomorrow.setDate(tomorrow.getDate() + 1);
                    tomorrow.setMinutes(0); // Optional: round to hour
                    const tomorrowEnd = new Date(tomorrow.getTime() + 60 * 60 * 1000); // +1 hour

                    document.getElementById('inputStartDt').value = this.toLocalISOString(tomorrow).slice(0, 16);
                    document.getElementById('inputEndDt').value = this.toLocalISOString(tomorrowEnd).slice(0, 16);
                    updateClockVisibility();
                });

                document.getElementById('inputStartDt').addEventListener('change', (e) => {
                    if (e.target.value) {
                        const sd = new Date(e.target.value);
                        const edInput = document.getElementById('inputEndDt');
                        const ed = edInput.value ? new Date(edInput.value) : null;
                        const isPlanMode = document.getElementById('btnModePlan').classList.contains('inactive') === false;

                        // Auto-set end time based on mode if not set or invalid
                        if (!ed || ed <= sd) {
                            const offset = isPlanMode ? (60 * 60 * 1000) : (5 * 60 * 1000);
                            const newEd = new Date(sd.getTime() + offset);
                            edInput.value = this.toLocalISOString(newEd).slice(0, 16);
                        }
                    }
                });

                btnSend.addEventListener('click', () => this.handleSend());

                document.getElementById('btnLocation').addEventListener('click', () => this.handleLocation());
                document.getElementById('btnVoice').addEventListener('click', () => this.handleVoice());
                document.getElementById('btnSync').addEventListener('click', () => this.requestSync());
                document.getElementById('btnDisconnect').addEventListener('click', () => this.disconnectGoogle());
                document.getElementById('btnImport').addEventListener('click', () => this.importFromGCal());

                inputText.addEventListener('keydown', (e) => {
                    if (e.key === 'Enter' && !e.shiftKey) {
                        e.preventDefault();
                        this.handleSend();
                    }
                });

                document.getElementById('btnEditSave').addEventListener('click', () => this.handleEditSave());
                document.getElementById('btnEditDelete').addEventListener('click', () => this.handleEditDelete());

                // Modals
                document.getElementById('locationModal').addEventListener('click', (e) => {
                    if (e.target.id === 'locationModal') this.closeLocationModal();
                });

                // Initialize default values for the first time
                setNowDefaults();
                updateClockVisibility();
                document.getElementById('editModal').addEventListener('click', (e) => {
                    if (e.target.id === 'editModal') this.closeEditModal();
                });
            },

            showToast: function (message) {
                const toast = document.getElementById('toast');
                toast.textContent = message;
                toast.classList.add('show');
                setTimeout(() => toast.classList.remove('show'), 3000);
            },

            showError: function (message) {
                const banner = document.getElementById('errorBanner');
                banner.textContent = message;
                banner.classList.add('show');
                setTimeout(() => banner.classList.remove('show'), 5000);
            },

            closeLocationModal: function () {
                document.getElementById('locationModal').classList.remove('active');
            },

            closeEditModal: function () {
                document.getElementById('editModal').classList.remove('active');
                this.currentEditId = null;
            },

            toLocalISOString: function (date) {
                const pad = n => n < 10 ? '0' + n : n;
                return date.getFullYear() + '-' +
                    pad(date.getMonth() + 1) + '-' +
                    pad(date.getDate()) + 'T' +
                    pad(date.getHours()) + ':' +
                    pad(date.getMinutes());
            },

            openEditModal: async function (id) {
                try {
                    const entries = await this.getAllEntries();
                    const entry = entries.find(e => e.id === id);
                    if (!entry) return;

                    this.currentEditId = id;

                    document.getElementById('editStartDt').value = this.toLocalISOString(new Date(entry.timestamp)).slice(0, 16);
                    document.getElementById('editEndDt').value = entry.endTimestamp ? this.toLocalISOString(new Date(entry.endTimestamp)).slice(0, 16) : '';
                    document.getElementById('editText').value = entry.content;

                    document.getElementById('editModal').classList.add('active');
                } catch (e) {
                    this.showError('繧ｨ繝ｳ繝医Μ縺ｮ隱ｭ縺ｿ霎ｼ縺ｿ縺ｫ螟ｱ謨励＠縺ｾ縺励◆');
                }
            },

            handleEditSave: async function () {
                if (!this.currentEditId) return;
                try {
                    const sv = document.getElementById('editStartDt').value;
                    const ev = document.getElementById('editEndDt').value;
                    const txt = document.getElementById('editText').value.trim();

                    if (!sv || !txt) {
                        this.showError('蠢・磯・岼縺悟・蜉帙＆繧後※縺・∪縺帙ｓ');
                        return;
                    }

                    const entries = await this.getAllEntries();
                    const entry = entries.find(e => e.id === this.currentEditId);
                    if (!entry) return;

                    entry.timestamp = new Date(sv).getTime();
                    entry.endTimestamp = ev ? new Date(ev).getTime() : null;
                    entry.content = txt;
                    entry.syncStatus = 'pending'; // mark modified

                    await this.saveEntry(entry);
                    this.closeEditModal();
                    await this.renderList();
                    this.updateSyncBadge();
                } catch (e) {
                    this.showError('菫晏ｭ倥↓螟ｱ謨励＠縺ｾ縺励◆: ' + e.message);
                }
            },

            handleEditDelete: async function () {
                if (!this.currentEditId) return;
                try {
                    const entries = await this.getAllEntries();
                    const entry = entries.find(e => e.id === this.currentEditId);
                    if (!entry) return;

                    if (entry.gcalId) {
                        entry.syncStatus = 'delete_pending';
                        await this.saveEntry(entry);
                    } else {
                        // physically delete if not synced to GCal yet
                        const transaction = this.db.transaction(['entries'], 'readwrite');
                        const store = transaction.objectStore('entries');
                        store.delete(entry.id);

                        // Wait for transaction to complete
                        await new Promise((resolve, reject) => {
                            transaction.oncomplete = () => resolve();
                            transaction.onerror = (e) => reject(e.target.error);
                        });
                    }
                    this.closeEditModal();
                    await this.renderList();
                    this.updateSyncBadge();
                } catch (e) {
                    this.showError('蜑企勁縺ｫ螟ｱ謨励＠縺ｾ縺励◆: ' + e.message);
                }
            },

            // --- Geolocation (Nominatim API) --- //
            handleLocation: function () {
                const modal = document.getElementById('locationModal');
                const list = document.getElementById('locationList');
                const loading = document.getElementById('locationLoading');

                modal.classList.add('active');
                list.innerHTML = '';
                loading.style.display = 'block';

                if (!navigator.geolocation) {
                    this.showError('迴ｾ蝨ｨ蝨ｰ讖溯・縺後し繝昴・繝医＆繧後※縺・∪縺帙ｓ');
                    this.closeLocationModal();
                    return;
                }

                navigator.geolocation.getCurrentPosition(
                    async (pos) => {
                        try {
                            const lat = pos.coords.latitude;
                            const lon = pos.coords.longitude;
                            // Nominatim API call
                            const url = `https://nominatim.openstreetmap.org/reverse?format=json&lat=${lat}&lon=${lon}&zoom=18&addressdetails=1&accept-language=ja`;
                            const res = await fetch(url, { headers: { 'User-Agent': 'QuickLog/1.0' } });
                            if (!res.ok) throw new Error('API Error');

                            const data = await res.json();
                            const addr = data.address || {};

                            const amenity = addr.amenity || addr.shop || addr.office || addr.tourism || addr.leisure || addr.building;
                            const road = addr.road;
                            const suburb = addr.suburb || addr.neighbourhood;
                            const city = addr.city || addr.town || addr.village || addr.county;

                            const candidates = new Set();

                            if (amenity) candidates.add(amenity);
                            if (amenity && road) candidates.add(`${amenity} (${road})`);
                            if (road && (suburb || city)) candidates.add(`${road} (${suburb || city})`);
                            if (suburb && city) candidates.add(`${suburb}, ${city}`);
                            else if (city) candidates.add(city);

                            if (candidates.size === 0) {
                                candidates.add(data.display_name.split(',')[0]); // fallback
                            }

                            loading.style.display = 'none';

                            Array.from(candidates).slice(0, 4).forEach(text => {
                                const btn = document.createElement('button');
                                btn.className = 'location-item';
                                btn.textContent = text;
                                btn.onclick = () => {
                                    const input = document.getElementById('inputText');
                                    input.value = input.value ? input.value + ' ' + text : text;
                                    input.dispatchEvent(new Event('input')); // trigger resize
                                    this.closeLocationModal();
                                    input.focus();
                                };
                                list.appendChild(btn);
                            });

                        } catch (e) {
                            loading.style.display = 'none';
                            this.showError('蝨ｰ轤ｹ諠・ｱ縺ｮ蜿門ｾ励↓螟ｱ謨励＠縺ｾ縺励◆');
                            this.closeLocationModal();
                        }
                    },
                    (err) => {
                        loading.style.display = 'none';
                        if (err.code === err.PERMISSION_DENIED) {
                            this.showError('菴咲ｽｮ諠・ｱ縺ｮ險ｱ蜿ｯ縺悟ｿ・ｦ√〒縺・);
                        } else if (err.code === err.TIMEOUT) {
                            this.showError('菴咲ｽｮ諠・ｱ縺ｮ蜿門ｾ励′繧ｿ繧､繝繧｢繧ｦ繝医＠縺ｾ縺励◆');
                        } else {
                            this.showError('迴ｾ蝨ｨ蝨ｰ縺ｮ蜿門ｾ励↓螟ｱ謨励＠縺ｾ縺励◆');
                        }
                        this.closeLocationModal();
                    },
                    { enableHighAccuracy: true, timeout: 10000, maximumAge: 0 }
                );
            },

            // --- Voice Input (Web Speech API) --- //
            handleVoice: function () {
                const btnVoice = document.getElementById('btnVoice');
                const inputText = document.getElementById('inputText');

                if (this.recognition && this.isRecording) {
                    this.recognition.stop();
                    return;
                }

                const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
                if (!SpeechRecognition) {
                    this.showError('髻ｳ螢ｰ蜈･蜉帙′繧ｵ繝昴・繝医＆繧後※縺・∪縺帙ｓ');
                    return;
                }

                if (!this.recognition) {
                    this.recognition = new SpeechRecognition();
                    this.recognition.lang = 'ja-JP';
                    this.recognition.interimResults = true;
                    this.recognition.continuous = false;

                    this.recognition.onstart = () => {
                        this.isRecording = true;
                        btnVoice.classList.add('recording');
                        btnVoice.textContent = '竢ｹ';
                        this.voiceStartText = inputText.value;
                    };

                    this.recognition.onresult = (event) => {
                        let interimTranscript = '';
                        let finalTranscript = '';

                        for (let i = event.resultIndex; i < event.results.length; ++i) {
                            if (event.results[i].isFinal) {
                                finalTranscript += event.results[i][0].transcript;
                            } else {
                                interimTranscript += event.results[i][0].transcript;
                            }
                        }

                        const currentText = this.voiceStartText + (this.voiceStartText && (finalTranscript || interimTranscript) ? ' ' : '');
                        inputText.value = currentText + finalTranscript + interimTranscript;
                        inputText.dispatchEvent(new Event('input')); // trigger resize

                        if (finalTranscript) {
                            this.voiceStartText = inputText.value;
                        }
                    };

                    this.recognition.onerror = (event) => {
                        if (event.error !== 'no-speech') {
                            this.showError('髻ｳ螢ｰ蜈･蜉帙お繝ｩ繝ｼ: ' + event.error);
                        }
                        this.stopRecording();
                    };

                    this.recognition.onend = () => {
                        this.stopRecording();
                    };
                }

                try {
                    this.recognition.start();
                } catch (e) { /* ignore if already started */ }
            },

            stopRecording: function () {
                this.isRecording = false;
                const btnVoice = document.getElementById('btnVoice');
                btnVoice.classList.remove('recording');
                btnVoice.textContent = '痔';
                document.getElementById('inputText').focus();
            },

            // --- Google Identity Services & Calendar Sync --- //
            initGIS: function () {
                if (typeof google === 'undefined' || !google.accounts) {
                    setTimeout(() => this.initGIS(), 200);
                    return;
                }

                this.tokenClient = google.accounts.oauth2.initTokenClient({
                    client_id: this.clientId,
                    scope: 'https://www.googleapis.com/auth/calendar.events',
                    callback: (tokenResponse) => {
                        if (tokenResponse && tokenResponse.access_token) {
                            this.accessToken = tokenResponse.access_token;
                            // GIS tokens are valid for 3600 seconds. We refresh 5 mins earlier.
                            this.tokenExpiry = Date.now() + ((tokenResponse.expires_in - 300) * 1000);
                            localStorage.setItem('ql_token', this.accessToken);
                            localStorage.setItem('ql_expiry', this.tokenExpiry);

                            this.updateAuthUI(true);
                            this.scheduleTokenRefresh();

                            if (this.pendingAuthAction === 'import') {
                                this.doImport();
                            } else {
                                this.doSync(); // Resume sync
                            }
                            this.pendingAuthAction = null;
                        } else {
                            this.showError('Google隱崎ｨｼ縺ｫ螟ｱ謨励＠縺ｾ縺励◆');
                        }
                    }
                });

                // Restore from local storage
                const savedToken = localStorage.getItem('ql_token');
                const savedExpiry = localStorage.getItem('ql_expiry');
                if (savedToken && savedExpiry && Number(savedExpiry) > Date.now()) {
                    this.accessToken = savedToken;
                    this.tokenExpiry = Number(savedExpiry);
                    this.updateAuthUI(true);
                    this.scheduleTokenRefresh();
                } else if (savedToken) {
                    // Expired or invalid
                    this.disconnectGoogle();
                }
            },

            updateAuthUI: function (isAuthenticated) {
                document.getElementById('btnDisconnect').style.display = isAuthenticated ? 'block' : 'none';
            },

            scheduleTokenRefresh: function () {
                if (this.refreshTimer) clearTimeout(this.refreshTimer);
                if (!this.tokenExpiry) return;

                const timeToRefresh = this.tokenExpiry - Date.now();
                if (timeToRefresh > 0) {
                    this.refreshTimer = setTimeout(() => {
                        if (this.tokenClient) {
                            this.tokenClient.requestAccessToken({ prompt: '' });
                        }
                    }, timeToRefresh);
                }
            },

            disconnectGoogle: function () {
                this.accessToken = null;
                this.tokenExpiry = null;
                localStorage.removeItem('ql_token');
                localStorage.removeItem('ql_expiry');
                if (this.refreshTimer) clearTimeout(this.refreshTimer);
                this.updateAuthUI(false);
                this.showToast('騾｣謳ｺ繧定ｧ｣髯､縺励∪縺励◆');
            },

            requestSync: async function () {
                if (!this.tokenClient) {
                    this.showError('Google騾｣謳ｺ縺ｮ貅門ｙ荳ｭ縺ｧ縺・);
                    return;
                }

                if (this.accessToken && this.tokenExpiry && this.tokenExpiry > Date.now()) {
                    // 蜷梧悄蜃ｦ逅・ｒ螳御ｺ・＆縺帙※縺九ｉGoogle繧ｫ繝ｬ繝ｳ繝繝ｼ繧帝幕縺・
                    this.showToast('繧ｫ繝ｬ繝ｳ繝繝ｼ縺ｫ蜷梧悄荳ｭ縺ｧ縺・..');
                    await this.doSync();
                    window.open('https://calendar.google.com/', '_blank');
                } else {
                    // 譛ｪ隱崎ｨｼ縺ｮ蝣ｴ蜷医・繝昴ャ繝励い繝・・遶ｶ蜷医ｒ驕ｿ縺代ｋ縺溘ａ繧ｫ繝ｬ繝ｳ繝繝ｼ繧帝幕縺九★縲∬ｪ崎ｨｼ繝輔Ο繝ｼ縺ｮ縺ｿ髢句ｧ九☆繧・
                    this.pendingAuthAction = 'sync';
                    this.tokenClient.requestAccessToken({ prompt: '' });
                }
            },

            doSync: async function () {
                try {
                    const entries = await this.getAllEntries();
                    const targets = entries.filter(e => e.syncStatus === 'pending' || e.syncStatus === 'delete_pending');

                    if (targets.length === 0) {
                        this.showToast('蜷梧悄貂医∩縺ｧ縺・笨・);
                        return;
                    }

                    let successCount = 0;
                    let failCount = 0;

                    for (const entry of targets) {
                        try {
                            if (entry.syncStatus === 'delete_pending' && entry.gcalId) {
                                await this.gcalApiRequest('DELETE', `/calendars/primary/events/${entry.gcalId}`);
                                // Remove physically from DB as per spec
                                const transaction = this.db.transaction(['entries'], 'readwrite');
                                const store = transaction.objectStore('entries');
                                store.delete(entry.id);
                                successCount++;
                            } else if (entry.syncStatus === 'pending') {
                                const eventBody = {
                                    summary: entry.content,
                                    start: { dateTime: new Date(entry.timestamp).toISOString() },
                                    end: { dateTime: new Date(entry.endTimestamp || (entry.timestamp + 60000)).toISOString() },
                                    extendedProperties: { private: { source: 'quicklog' } }
                                };
                                if (entry.colorId) eventBody.colorId = entry.colorId;

                                if (entry.gcalId) {
                                    // Update
                                    await this.gcalApiRequest('PUT', `/calendars/primary/events/${entry.gcalId}`, eventBody);
                                } else {
                                    // Create
                                    const res = await this.gcalApiRequest('POST', '/calendars/primary/events', eventBody);
                                    entry.gcalId = res.id;
                                }
                                entry.syncStatus = 'synced';
                                await this.saveEntry(entry);
                                successCount++;
                            }
                        } catch (err) {
                            console.error('Item sync failed', entry, err);
                            failCount++;
                            if (err.status === 401) {
                                this.disconnectGoogle(); // Invalid token
                                throw err;
                            }
                        }
                    }

                    if (failCount > 0) {
                        this.showError(`${failCount}莉ｶ縺ｮ繧ｨ繝ｩ繝ｼ縺檎匱逕溘＠縺ｾ縺励◆`);
                    } else if (successCount > 0) {
                        this.showToast(`${successCount}莉ｶ蜷梧悄縺励∪縺励◆ 笨伝);
                    }

                    await this.renderList();
                    this.updateSyncBadge();
                } catch (e) {
                    if (e.status !== 401) {
                        this.showError('蜷梧悄蜃ｦ逅・〒繧ｨ繝ｩ繝ｼ縺檎匱逕溘＠縺ｾ縺励◆');
                    }
                }
            },

            importFromGCal: function () {
                if (!this.tokenClient) {
                    this.showError('Google騾｣謳ｺ縺ｮ貅門ｙ荳ｭ縺ｧ縺・);
                    return;
                }

                if (this.accessToken && this.tokenExpiry && this.tokenExpiry > Date.now()) {
                    this.doImport();
                } else {
                    this.pendingAuthAction = 'import';
                    this.tokenClient.requestAccessToken({ prompt: '' });
                }
            },

            doImport: async function () {
                const btn = document.getElementById('btnImport');
                btn.disabled = true;
                btn.textContent = '隱ｭ縺ｿ霎ｼ縺ｿ荳ｭ...';

                try {
                    const timeMin = new Date();
                    timeMin.setDate(timeMin.getDate() - 90);
                    const timeMax = new Date();
                    timeMax.setDate(timeMax.getDate() + 60);

                    const res = await this.gcalApiRequest('GET',
                        `/calendars/primary/events?timeMin=${encodeURIComponent(timeMin.toISOString())}&timeMax=${encodeURIComponent(timeMax.toISOString())}&maxResults=500&singleEvents=true&orderBy=startTime`
                    );

                    const existingEntries = await this.getAllEntries();
                    const existingGcalIds = new Set(existingEntries.map(e => e.gcalId).filter(id => id));

                    let imported = 0;
                    for (const item of (res.items || [])) {
                        if (existingGcalIds.has(item.id)) continue;

                        const startTs = new Date(item.start.dateTime || item.start.date).getTime();
                        const endTs = item.end ? new Date(item.end.dateTime || item.end.date).getTime() : null;

                        const newEntry = {
                            id: `${startTs}_${Math.random().toString(36).substr(2, 9)}`,
                            content: item.summary || '',
                            timestamp: startTs,
                            endTimestamp: endTs,
                            gcalId: item.id,
                            colorId: item.colorId || null,
                            syncStatus: 'synced'
                        };
                        await this.saveEntry(newEntry);
                        imported++;
                    }

                    this.showToast(`${imported}莉ｶ隱ｭ縺ｿ霎ｼ縺ｿ縺ｾ縺励◆`);
                    await this.renderList();
                    this.updateSyncBadge();
                } catch (e) {
                    if (e.status === 401) {
                        this.disconnectGoogle();
                    } else {
                        this.showError('隱ｭ縺ｿ霎ｼ縺ｿ縺ｫ螟ｱ謨励＠縺ｾ縺励◆');
                    }
                } finally {
                    btn.disabled = false;
                    btn.textContent = '驕主悉縺ｮ險倬鹸繧竪oogle繧ｫ繝ｬ繝ｳ繝繝ｼ縺九ｉ隱ｭ縺ｿ霎ｼ繧';

                    const existingEntries = await this.getAllEntries();
                    const hasImported = existingEntries.some(e => e.gcalId && e.syncStatus === 'synced');
                    document.getElementById('importSection').style.display = hasImported ? 'none' : 'block';
                }
            },

            gcalApiRequest: async function (method, path, body = null) {
                const headers = {
                    'Authorization': `Bearer ${this.accessToken}`,
                    'Content-Type': 'application/json'
                };

                const conf = { method, headers };
                if (body) conf.body = JSON.stringify(body);

