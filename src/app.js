// --- 1. Setup & Helpers ---
        let rows = [];
        let currentTemplate = 1;
        let logoImage = null; 
        let darkLogoImage = null;
        let lastBlob = null; 
        let lastZipBlob = null; 
        let lastScreenSignature = null;
        let lastReportSignature = null; 

        const FIXED_LOGO_URL = "./public/assets/logo-footer.png";
        const DARK_LOGO_URL = "./public/assets/logo-footer-dark.png";
        const FALLBACK_LOGO_URL = "https://nauss.edu.sa/Style%20Library/ar-sa/Styles/images/home/logo-footer.png";
        
        const UNI = {
            green: "#016564",
            gold: "#d0b284",
            white: "#ffffff",
            grayLight: "#f8f9f9",
            grayMid: "#d6d7d4",
            grayDark: "#98aaaa",
            greenMid: "#498983",
            supRed: "#802f2d",
            iosGray: "#8e8e93",
            iosBg: "#f5f5f7",
            paperBg: "#f4e4bc",
            paperPattern: "rgba(161, 136, 127, 0.15)",
            bodyBg: "#f7f5f5"
        };

        const COLORS = { paperColors: ["#fff9c4", "#e3f2fd", "#fce4ec", "#e8f5e9", "#fff3e0", "#f3e5f5"] };

        const DDL = {
            sup: ["نايف الشهراني", "عبدالمحسن العنزي", "خازم الأسمري", "وليد السليم", "ثامر العاصمي", "فؤاد نعمان", "مزنة آل محمود", "بشاير الطحيني", "يعقوب الشمري", "نواف المحارب"],
            room: ["LAB 1", "LAB 2", "LAB 3", "LAB 4", "LAB 5", "LAB 6", "CLASS 1", "CLASS 2", "CLASS 3", "CLASS 4", "CLASS 5", "CLASS 6", "CLASS 7", "CLASS 8", "CLASS 9", "CLASS 10", "مركز الجرائم السيبرانية", "مركز الذكاء الاصطناعي", "مركز السلامة المرورية", "كلية الادلة الجنائية", "النادي الرياضي", "خارج المملكة", "خارج الرياض", "مركز التدريب - الامن العام"],
            floor: ["الأرضي", "الأول", "الثاني", "خارجي", "عن بعد"],
            period: ["صباحي", "مسائي"],
            status: ["جديدة", "مستمرة", "خارجية", "عن بعد"],
            type: ["داخلية", "دولية"]
        };

        const KEYS = { n: "اسم النشاط التدريبي", loc: "مكان التنفيذ", s: "تاريخ البدء", e: "تاريخ الانتهاء", st: "الحالة", sp: "اسم منسق التدريب", r: "القاعة", f: "الطابق", p: "الفترة" };
        const INTERNAL_EXECUTION_KEYWORDS = ["السعودية", "المملكة", "الرياض", "جدة", "الدمام", "الجامعة", "جامعة نايف", "نايف"];
        const DEFAULT_EXECUTION_PLACE = "مقر الجامعة";
        const SCREEN_EMAIL_RECIPIENTS = ["KAhmad@nauss.edu.sa", "AMohammad@nauss.edu.sa", "WAlsaleem@nauss.edu.sa"];
        const AGENT_REPORT_RECIPIENTS = ["T-AAlmargan@nauss.edu.sa"];
        const ARCH_KEY = "nfdp_archive_v42";
        const MEMORY_KEY = "nfdp_memory_v1";
        const INTRO_KEY = "nfdp_intro_settings_v1";
        const DEFAULT_SCREEN_INTRO = "ينفذ برنامج الشراكات الدولية بوزارة الداخلية بالتعاون مع جامعة نايف العربية للعلوم الأمنية البرامج التدريبية التالية:";
        let lastUsed = {};

        function showToast(message, type = 'info') { const container = document.getElementById('toast-container'); const toast = document.createElement('div'); toast.className = `toast ${type}`; toast.innerText = message; container.appendChild(toast); setTimeout(() => { toast.style.animation = 'fadeOut 0.3s ease-out forwards'; setTimeout(() => toast.remove(), 300); }, 3000); }

        function parseExcelDate(value) {
            if (!value) return '';
            const formatDate = (dateObj) => {
                let year = dateObj.getFullYear();
                let month = (dateObj.getMonth() + 1).toString().padStart(2, '0');
                let day = dateObj.getDate().toString().padStart(2, '0');
                return `${year}-${month}-${day}`;
            };

            if (value instanceof Date) return formatDate(value);

            if (typeof value === 'number') {
                const date = new Date((value - 25569) * 86400 * 1000);
                return formatDate(date);
            }

            if (typeof value === 'string') {
                const raw = value.trim();
                if (!raw) return '';

                const normalized = raw
                    .replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d))
                    .replace(/[۰-۹]/g, d => '۰۱۲۳۴۵۶۷۸۹'.indexOf(d))
                    .replace(/[\/\.]/g, '-');

                let match = normalized.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
                if (match) {
                    const y = match[1];
                    const m = match[2].padStart(2, '0');
                    const d = match[3].padStart(2, '0');
                    return `${y}-${m}-${d}`;
                }

                match = normalized.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
                if (match) {
                    const d = match[1].padStart(2, '0');
                    const m = match[2].padStart(2, '0');
                    const y = match[3];
                    return `${y}-${m}-${d}`;
                }

                const date = new Date(raw);
                if (!isNaN(date.getTime())) return formatDate(date);
            }

            return '';
        }
        
        // Helper to format date nicely (YYYY/MM/DD)
        function formatDateShort(dateStr) {
            if(!dateStr) return '';
            // Replace - with /
            return dateStr.replace(/-/g, '/');
        }

        function normalizeArabicValue(value) {
            return String(value || '')
                .trim()
                .replace(/[إأآا]/g, 'ا')
                .replace(/ى/g, 'ي')
                .replace(/ة/g, 'ه')
                .replace(/\s+/g, ' ')
                .toLowerCase();
        }

        function normalizeExecutionPlace(value) {
            const raw = String(value || '').trim();
            const v = normalizeArabicValue(raw);
            if (!v) return DEFAULT_EXECUTION_PLACE;
            if (v.includes('جامعه نايف') || v.includes('نايف العربيه') || v.includes('مقر الجامعه') || v.includes('مقر الجامعة')) return 'مقر الجامعة';
            return raw;
        }

        function isInternalExecutionPlace(value) {
            const v = normalizeArabicValue(value);
            if (!v) return true;
            return v.includes('جامعه نايف') || v.includes('جامعة نايف') || v.includes('نايف العربيه') || v.includes('مقر الجامعه') || v.includes('مقر الجامعة') || v === normalizeArabicValue(DEFAULT_EXECUTION_PLACE);
        }

        function applyExecutionPlaceRules(row) {
            if (!row) return row;
            row.loc = normalizeExecutionPlace(row.loc);
            if (!isInternalExecutionPlace(row.loc)) {
                row.f = 'خارجي';
                row.r = 'خارج المملكة';
            }
            return row;
        }


        function extractTrainingRoom(value) {
            const raw = String(value || '').trim();
            if (!raw) return '';
            const normalized = raw.replace(/\s+/g, ' ');
            const labMatch = normalized.match(/LAB\s*-?\s*(\d+)/i);
            if (labMatch) return `LAB ${labMatch[1]}`;
            const classMatch = normalized.match(/CLASS\s*-?\s*(\d+)/i);
            if (classMatch) return `CLASS ${classMatch[1]}`;
            return raw.replace(/^برج التدريب\s*-\s*/i, '').trim();
        }

        function inferFloorFromRoom(room, fallback = '') {
            const r = String(room || '').toUpperCase();
            if (r.includes('LAB')) return 'الثاني';
            if (r.includes('CLASS')) return 'الأول';
            return fallback || '';
        }

        function isLmsDetailLine(parts) {
            const first = normalizeArabicValue(parts[0] || '');
            return first.startsWith('المجال التدريبي') ||
                   first.startsWith('معرف التدريب') ||
                   first.startsWith('التصنيف') ||
                   first.startsWith('الوقت المفضل');
        }

        function extractPreferredPeriodFromText(line) {
            const match = String(line || '').match(/الوقت\s*المفضل\s*:\s*([^\t|,;]+)/i);
            return match ? normalizePeriod(match[1].trim()) : '';
        }

        function shouldHideFloorValue(value) {
            const v = normalizeArabicValue(value);
            return v === 'خارجي' || v === 'خارج الرياض' || v === 'خارج المملكه' || v === 'خارج المملكة';
        }


        function normalizePeriod(value) {
            const v = normalizeArabicValue(value);
            if (['ص', '1', 'الصبح', 'صباح', 'صباحي'].includes(v)) return 'صباحي';
            if (['م', '2', 'مساء', 'المساء', 'مساءي', 'مسائي'].includes(v)) return 'مسائي';
            return String(value || DDL.period[0]).trim();
        }

        function normalizeFloor(value) {
            const v = normalizeArabicValue(value).replace(/ /g, '');
            if (['1', 'الاول', 'اول'].includes(v)) return 'الأول';
            if (['2', 'الثاني', 'ثاني'].includes(v)) return 'الثاني';
            if (['3', 'الثالث', 'ثالث'].includes(v)) return 'الثالث';
            if (['0', 'ارضي', 'الارضي'].includes(v)) return 'الأرضي';
            return String(value || DDL.floor[0]).trim();
        }

        function normalizeStatus(value) {
            const v = normalizeArabicValue(value);
            if (v === 'مؤكد' || v === 'موكد') return 'جديدة';
            if (v === 'قيد التقدم') return 'مستمرة';
            return String(value || DDL.status[0]).trim();
        }

        function getWeekNumberFromDate(dateStr) {
            const date = new Date(dateStr);
            if (isNaN(date.getTime())) return '';
            const temp = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
            const dayNum = temp.getUTCDay() || 7;
            temp.setUTCDate(temp.getUTCDate() + 4 - dayNum);
            const yearStart = new Date(Date.UTC(temp.getUTCFullYear(), 0, 1));
            return Math.ceil((((temp - yearStart) / 86400000) + 1) / 7);
        }

        function updateWeekNumberFromRows() {
            const dates = rows.map(r => r.s).filter(Boolean).sort();
            const week = dates.length ? getWeekNumberFromDate(dates[0]) : getWeekNumberFromDate(new Date().toISOString().slice(0,10));
            if (week) document.getElementById('w-id').value = week;
        }

        function getSelectedOwner() {
            return document.getElementById('owner-select').value || '';
        }

        function validateOwnerSelection() {
            if (!getSelectedOwner()) {
                showToast('اختر الموظف المسؤول أولاً', 'error');
                document.getElementById('owner-select').focus();
                return false;
            }
            return true;
        }

        try { const fontCairo = new FontFace('Cairo', 'url(https://fonts.gstatic.com/s/cairo/v28/SLXgc1nY6HkvangtZmpQdkhzfH5lkSs2SgRjCAGMQ1z0hGA-W1M.woff2)'); document.fonts.add(fontCairo); fontCairo.load(); } catch(e) {}

        function loadDefaultLogo() {
            const img = new Image();
            img.crossOrigin = "Anonymous";
            img.onload = () => { logoImage = img; };
            img.onerror = () => {
                const fallback = new Image();
                fallback.crossOrigin = "Anonymous";
                fallback.onload = () => { logoImage = fallback; };
                fallback.src = FALLBACK_LOGO_URL;
            };
            img.src = FIXED_LOGO_URL;
        }

        function loadDarkLogo() {
            const img = new Image();
            img.crossOrigin = "Anonymous";
            img.onload = () => { darkLogoImage = img; };
            img.onerror = () => { darkLogoImage = null; };
            img.src = DARK_LOGO_URL;
        }

        document.addEventListener('DOMContentLoaded', () => { addRow(); initWeekNumber(); loadDefaultLogo(); loadDarkLogo(); loadMemory(); loadIntroSettings(); setupSmartPasteUI(); bindStateInvalidationEvents(); checkShareSupport(); updateShareButtons(); });

        // --- 2. UI Logic ---
        function initWeekNumber() { updateWeekNumberFromRows(); }

        function showPage(pageId) {
            document.querySelectorAll('.page-container').forEach(el => el.classList.remove('active'));
            const target = document.getElementById('page-' + pageId);
            if (target) target.classList.add('active');
            document.querySelectorAll('.nav-tab').forEach(el => el.classList.remove('active'));
            const tabs = document.querySelectorAll('.nav-tab');
            if (pageId === 'archive' && tabs[1]) tabs[1].classList.add('active');
            if (pageId === 'generator' && tabs[0]) tabs[0].classList.add('active');
            if(pageId === 'archive') refreshArchiveView();
        }

        function resetAndShowGenerator() {
            if(rows.length > 1 && !confirm("هل تريد بدء جدول جديد؟")) return;
            clearCurrentData();
            showPage('generator');
        }

        function clearCurrentData() {
            rows = [];
            const box = document.getElementById('rows-box');
            if (box) box.innerHTML = '';
            addRow();
            updateIdx();
            updateWeekNumberFromRows();
            showToast("تم مسح البيانات", 'info');
        }

        function selectTemplate(id) {
            currentTemplate = id;
            document.querySelectorAll('.template-thumb').forEach(el => el.classList.remove('selected'));
            const selected = document.getElementById('tpl-' + id);
            if (selected) selected.classList.add('selected');
            invalidateGeneratedFiles();
        }

        function toggleLoader(show, text = "جاري المعالجة...") {
            const el = document.getElementById('loader');
            const txt = document.getElementById('loader-text');
            if (!el) return;
            if (txt) txt.innerText = text;
            if(show) { el.classList.remove('hidden'); el.classList.add('flex'); }
            else { el.classList.add('hidden'); el.classList.remove('flex'); }
        }

        function loadMemory() {
            try { const m = localStorage.getItem(MEMORY_KEY); if(m) lastUsed = JSON.parse(m); }
            catch(e) { lastUsed = {}; }
        }

        function saveMemory(key, val) {
            try { lastUsed[key] = val; localStorage.setItem(MEMORY_KEY, JSON.stringify(lastUsed)); }
            catch(e) {}
        }

        function loadIntroSettings() {
            try {
                const saved = JSON.parse(localStorage.getItem(INTRO_KEY) || '{}');
                const modeEl = document.getElementById('intro-mode');
                const customEl = document.getElementById('intro-custom');
                if (!modeEl || !customEl) return;
                modeEl.value = saved.mode || 'default';
                customEl.value = saved.customText || '';
                handleIntroModeChange(false);
            } catch(e) {}
        }

        function saveIntroSettings() {
            const modeEl = document.getElementById('intro-mode');
            const customEl = document.getElementById('intro-custom');
            if (!modeEl || !customEl) return;
            localStorage.setItem(INTRO_KEY, JSON.stringify({ mode: modeEl.value, customText: customEl.value }));
        }

        function handleIntroModeChange(showMessage = true) {
            const modeEl = document.getElementById('intro-mode');
            const customEl = document.getElementById('intro-custom');
            if (!modeEl || !customEl) return;
            customEl.classList.toggle('hidden', modeEl.value !== 'custom');
            saveIntroSettings();
            if (showMessage) {
                invalidateGeneratedFiles();
                showToast('تم تحديث إعداد النص التمهيدي', 'success');
            }
        }

        function getScreenIntroText() {
            const modeEl = document.getElementById('intro-mode');
            const customEl = document.getElementById('intro-custom');
            const mode = modeEl ? modeEl.value : 'default';
            if (mode === 'none') return '';
            if (mode === 'custom') return (customEl ? customEl.value.trim() : '');
            return DEFAULT_SCREEN_INTRO;
        }

        function getRowsSignaturePayload() {
            return rows
                .filter(x => String(x.n || '').trim() !== '')
                .map(x => ({
                    n: x.n || '',
                    loc: x.loc || '',
                    s: x.s || '',
                    e: x.e || '',
                    st: x.st || '',
                    p: x.p || '',
                    f: x.f || '',
                    r: x.r || '',
                    sp: x.sp || ''
                }));
        }

        function getWeekValue() {
            const weekEl = document.getElementById('w-id');
            return weekEl ? String(weekEl.value || '').trim() : '';
        }

        function getScreenGenerationSignature() {
            return JSON.stringify({
                output: 'screen',
                week: getWeekValue(),
                owner: getSelectedOwner(),
                template: currentTemplate,
                intro: getScreenIntroText(),
                rows: getRowsSignaturePayload()
            });
        }

        function getAgentReportGenerationSignature() {
            return JSON.stringify({
                output: 'agent-report',
                week: getWeekValue(),
                owner: getSelectedOwner(),
                rows: getRowsSignaturePayload()
            });
        }

        function isScreenShareReady() {
            return !!lastZipBlob && lastScreenSignature === getScreenGenerationSignature();
        }

        function isAgentReportShareReady() {
            return !!lastBlob && lastReportSignature === getAgentReportGenerationSignature();
        }

        function updateShareButtonState(button, isReady, readyTitle) {
            if (!button) return;
            button.classList.remove('hidden');
            button.title = isReady ? readyTitle : 'لا يمكن مشاركة الملف إلا بعد توليد الملف بشكله النهائي';
            button.classList.toggle('opacity-50', !isReady);
            button.classList.toggle('cursor-not-allowed', !isReady);
        }

        function updateShareButtons() {
            updateShareButtonState(document.getElementById('btn-share-screen'), isScreenShareReady(), 'تنزيل مسودة بريد شاشة مدخل المبنى');
            updateShareButtonState(document.getElementById('btn-share-report'), isAgentReportShareReady(), 'تنزيل مسودة بريد الوكيل');
        }

        function invalidateGeneratedFiles() {
            lastZipBlob = null;
            lastBlob = null;
            lastScreenSignature = null;
            lastReportSignature = null;
            updateShareButtons();
        }

        function bindStateInvalidationEvents() {
            const ownerEl = document.getElementById('owner-select');
            const weekEl = document.getElementById('w-id');
            const customIntroEl = document.getElementById('intro-custom');
            if (ownerEl) ownerEl.addEventListener('change', invalidateGeneratedFiles);
            if (weekEl) weekEl.addEventListener('input', invalidateGeneratedFiles);
            if (customIntroEl) customIntroEl.addEventListener('input', invalidateGeneratedFiles);
        }

        // الشعار الرسمي ثابت ويحمّل تلقائيًا من public/assets/logo-footer.png

        function normalizeLegacyRow(d = {}) {
            return applyExecutionPlaceRules({
                id: d.id || Date.now() + Math.random(),
                n: d.n || d[KEYS.n] || "",
                loc: d.loc || d.place || d.executionPlace || d[KEYS.loc] || lastUsed.loc || DEFAULT_EXECUTION_PLACE,
                s: d.s || d[KEYS.s] || "",
                e: d.e || d[KEYS.e] || "",
                st: normalizeStatus(d.st || d[KEYS.st] || lastUsed.st || DDL.status[0]),
                sp: d.sp || d[KEYS.sp] || lastUsed.sp || DDL.sup[0],
                r: extractTrainingRoom(d.r || d[KEYS.r] || lastUsed.r || DDL.room[0]),
                f: normalizeFloor(d.f || d[KEYS.f] || inferFloorFromRoom(d.r || d[KEYS.r], lastUsed.f || DDL.floor[0])),
                p: normalizePeriod(d.p || d[KEYS.p] || lastUsed.p || DDL.period[0])
            });
        }

        function addRow(d = {}) { 
            const o = normalizeLegacyRow(d);
            if (!o.id) o.id = Date.now() + Math.random();
            rows.push(o); renderRow(o); 
            invalidateGeneratedFiles();
        }
        
        function renderRow(o) {
            const el = document.createElement('div');
            el.id = `r-${o.id}`;
            el.className = "bg-white p-2 rounded-lg border hover:border-[#2c6060] transition-colors relative group";
            const opts = (arr, sel) => arr.map(x => `<option value="${x}" ${x===sel?'selected':''}>${x}</option>`).join('');
            el.innerHTML = `
                <div class="grid grid-cols-12 gap-2 items-center">
                    <div class="col-span-1 text-center text-slate-300 font-bold text-xs idx">${rows.length}</div>
                    <div class="col-span-11 md:col-span-4"><input type="text" placeholder="اسم النشاط التدريبي" value="${o.n}" onchange="upd(${o.id},'n',this.value)" class="form-input border-transparent focus:border-[#2c6060]"></div>
                    <div class="col-span-6 md:col-span-2"><input type="text" placeholder="مكان التنفيذ" value="${o.loc}" onchange="upd(${o.id},'loc',this.value)" class="form-input text-xs"></div>
                    <div class="col-span-6 md:col-span-2"><select onchange="upd(${o.id},'st',this.value)" class="form-input text-xs">${opts(DDL.status, o.st)}</select></div>
                    <div class="col-span-11 md:col-span-2"><select onchange="upd(${o.id},'sp',this.value)" class="form-input text-xs">${opts(DDL.sup, o.sp)}</select></div>
                    <div class="col-span-1 flex flex-col gap-1 items-center"><div class="w-6 h-6 flex items-center justify-center cursor-pointer bg-red-50 text-red-500 rounded hover:bg-red-100 text-sm font-bold" onclick="delRow(${o.id})">×</div></div>
                </div>
                <div class="grid grid-cols-12 gap-2 items-center mt-1">
                    <div class="hidden md:block col-span-1"></div>
                    <div class="col-span-6 md:col-span-2"><input type="date" value="${o.s}" onchange="upd(${o.id},'s',this.value)" class="form-input text-xs" title="تاريخ البدء"></div>
                    <div class="col-span-6 md:col-span-2"><input type="date" value="${o.e}" onchange="upd(${o.id},'e',this.value)" class="form-input text-xs" title="تاريخ الانتهاء"></div>
                    <div class="col-span-6 md:col-span-2"><select onchange="upd(${o.id},'r',this.value)" class="form-input text-xs">${opts(DDL.room, o.r)}</select></div>
                    <div class="col-span-6 md:col-span-2"><select onchange="upd(${o.id},'f',this.value)" class="form-input text-xs">${opts(DDL.floor, o.f)}</select></div>
                    <div class="col-span-12 md:col-span-2"><select onchange="upd(${o.id},'p',this.value)" class="form-input text-xs">${opts(DDL.period, o.p)}</select></div>
                </div>`;
            document.getElementById('rows-box').appendChild(el);
        }
        
        function upd(id, k, v) {
            let o = rows.find(x => x.id === id);
            if(o) {
                if (k === 'loc') v = normalizeExecutionPlace(v);
                if (k === 'p') v = normalizePeriod(v);
                if (k === 'f') v = normalizeFloor(v);
                if (k === 'st') v = normalizeStatus(v);
                o[k] = v;
                if (k === 'loc') {
                    applyExecutionPlaceRules(o);
                    const rowEl = document.getElementById(`r-${id}`);
                    if (rowEl) {
                        const index = rows.findIndex(x => x.id === id);
                        rowEl.remove();
                        renderRow(o);
                        const box = document.getElementById('rows-box');
                        const newEl = document.getElementById(`r-${id}`);
                        if (box && newEl && index >= 0 && index < box.children.length - 1) {
                            box.insertBefore(newEl, box.children[index]);
                        }
                        updateIdx();
                    }
                }
                if (k === 's') updateWeekNumberFromRows();
            }
            saveMemory(k, v);
            invalidateGeneratedFiles();
        }
        function delRow(id) { rows = rows.filter(x => x.id !== id); document.getElementById(`r-${id}`).remove(); updateIdx(); if (rows.length === 0) addRow(); invalidateGeneratedFiles(); }
        function updateIdx() { document.querySelectorAll('.idx').forEach((n, i) => n.innerText = i+1); }

        function isExternalExecution(place = "") {
            const value = String(place || "").trim();
            if (!value) return false;
            return !INTERNAL_EXECUTION_KEYWORDS.some(keyword => value.includes(keyword));
        }


        // --- 2.1 Smart Paste from LMS ---
        function setupSmartPasteUI() {
            const uploadLabel = document.querySelector('label[for="x-f"]') || document.getElementById('x-f')?.closest('label');
            if (!uploadLabel || document.getElementById('smart-paste-btn')) return;

            const holder = uploadLabel.parentElement?.parentElement?.parentElement || uploadLabel.parentElement;
            if (!holder) return;

            const btnWrap = document.createElement('div');
            btnWrap.className = 'mt-2';
            btnWrap.innerHTML = `
                <button id="smart-paste-btn" type="button" onclick="openSmartPasteModal()" class="form-input text-center text-xs hover:bg-slate-50">
                    📋 لصق ذكي من LMS
                </button>
            `;
            holder.appendChild(btnWrap);

            const modal = document.createElement('div');
            modal.id = 'smart-paste-modal';
            modal.className = 'fixed inset-0 bg-slate-900/50 z-[120] hidden items-center justify-center p-4';
            modal.innerHTML = `
                <div class="bg-white rounded-2xl shadow-xl w-full max-w-4xl border overflow-hidden">
                    <div class="flex items-center justify-between p-4 border-b">
                        <h3 class="font-bold text-[#2c6060]">اللصق الذكي من LMS</h3>
                        <button type="button" onclick="closeSmartPasteModal()" class="text-slate-400 hover:text-red-500 text-2xl leading-none">×</button>
                    </div>
                    <div class="p-4 space-y-3">
                        <p class="text-xs text-slate-500">
                            الصق الصفوف المنسوخة من منصة LMS. سيقرأ النظام الأعمدة بالترتيب:
                            اسم النشاط التدريبي | مكان التنفيذ | تاريخ البدء | تاريخ الانتهاء | الحالة | اسم منسق التدريب | القاعة | الطابق | الفترة
                        </p>
                        <textarea id="smart-paste-text" class="form-input min-h-[260px] font-mono text-sm leading-7" placeholder="الصق البيانات هنا..."></textarea>
                        <div class="flex flex-wrap gap-2 justify-between items-center">
                            <div class="text-xs text-slate-400" id="smart-paste-hint">يدعم النسخ من الجداول، Tab، والفواصل.</div>
                            <div class="flex gap-2">
                                <button type="button" onclick="clearSmartPasteText()" class="btn-outline text-xs">مسح</button>
                                <button type="button" onclick="applySmartPaste('append')" class="btn-outline text-xs">إضافة للجدول الحالي</button>
                                <button type="button" onclick="applySmartPaste('replace')" class="btn-main text-xs">استبدال الجدول الحالي</button>
                            </div>
                        </div>
                    </div>
                </div>
            `;
            document.body.appendChild(modal);
        }

        function openSmartPasteModal() {
            const modal = document.getElementById('smart-paste-modal');
            if (!modal) return;
            modal.classList.remove('hidden');
            modal.classList.add('flex');
            setTimeout(() => document.getElementById('smart-paste-text')?.focus(), 50);
        }

        function closeSmartPasteModal() {
            const modal = document.getElementById('smart-paste-modal');
            if (!modal) return;
            modal.classList.add('hidden');
            modal.classList.remove('flex');
        }

        function clearSmartPasteText() {
            const txt = document.getElementById('smart-paste-text');
            if (txt) txt.value = '';
        }

        function splitSmartPasteLine(line) {
            if (line.includes('\t')) return line.split('\t');
            if (line.includes('|')) return line.split('|');
            if (line.includes(';')) return line.split(';');
            if (line.includes(',')) return line.split(',');
            return line.split(/\s{2,}/);
        }

        function isSmartPasteHeader(parts) {
            const joined = normalizeArabicValue(parts.join(' '));
            return joined.includes('اسم النشاط') || joined.includes('مكان التنفيذ') || joined.includes('تاريخ البدء') || joined.includes('اسم منسق');
        }

        function buildSmartPasteRow(parts) {
            const clean = parts.map(v => String(v || '').trim());
            if (clean.length < 6 || isSmartPasteHeader(clean) || isLmsDetailLine(clean)) return null;

            const startDate = parseExcelDate(clean[2]) || clean[2] || '';
            const endDate = parseExcelDate(clean[3]) || clean[3] || '';
            if (!startDate || !endDate) return null;

            const loc = normalizeExecutionPlace(clean[1]);
            const rawRoom = clean[6] || '';
            const room = extractTrainingRoom(rawRoom);
            const floor = clean[7] ? normalizeFloor(clean[7]) : inferFloorFromRoom(room, DDL.floor[0]);
            const period = clean[8] ? normalizePeriod(clean[8]) : DDL.period[0];

            return applyExecutionPlaceRules({
                n: clean[0] || '',
                loc,
                s: startDate,
                e: endDate,
                st: normalizeStatus(clean[4]),
                sp: clean[5] || DDL.sup[0],
                r: room || DDL.room[0],
                f: floor || DDL.floor[0],
                p: period
            });
        }

        function parseSmartPasteText(rawText) {
            const raw = String(rawText || '').replace(/\r/g, '').trim();
            if (!raw) return [];

            const lines = raw.split('\n').map(line => line.trim()).filter(Boolean);
            const parsed = [];

            for (const line of lines) {
                const parts = splitSmartPasteLine(line).map(x => x.trim());

                const preferredPeriod = extractPreferredPeriodFromText(line);
                if (preferredPeriod && parsed.length) {
                    parsed[parsed.length - 1].p = preferredPeriod;
                    continue;
                }

                if (isSmartPasteHeader(parts) || isLmsDetailLine(parts)) continue;

                const row = buildSmartPasteRow(parts);
                if (row && row.n) parsed.push(row);
            }

            return parsed;
        }

        function applySmartPaste(mode = 'replace') {
            const txt = document.getElementById('smart-paste-text');
            const pastedRows = parseSmartPasteText(txt?.value || '');

            if (!pastedRows.length) {
                showToast('لم يتم العثور على بيانات قابلة للقراءة', 'error');
                return;
            }

            if (mode === 'replace') {
                rows = [];
                const box = document.getElementById('rows-box');
                if (box) box.innerHTML = '';
            }

            pastedRows.forEach(row => addRow(row));
            updateIdx();
            updateWeekNumberFromRows();
            invalidateGeneratedFiles();
            closeSmartPasteModal();
            showToast(`تمت قراءة ${pastedRows.length} نشاط تدريبي`, 'success');
        }


        // --- 3. Excel ---
        function dlTpl() {
            const sample = ["إدارة الحشود وتأمين الفعاليات الكبرى", "مقر الجامعة", "2026-04-19", "2026-04-23", "جديدة", "نايف الشهراني", "LAB 1", "الثاني", "صباحي"];
            const ws = XLSX.utils.aoa_to_sheet([[KEYS.n, KEYS.loc, KEYS.s, KEYS.e, KEYS.st, KEYS.sp, KEYS.r, KEYS.f, KEYS.p], sample]);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "Template");
            XLSX.writeFile(wb, "NAUSS_Weekly_Schedule_Template.xlsx");
            showToast("تم تحميل القالب", 'success');
        }
        function rdExcel(e) {
            const f = e.target.files[0]; if(!f) return;
            const rd = new FileReader();
            rd.onload = function(ev) {
                try {
                    const data = new Uint8Array(ev.target.result);
                    const workbook = XLSX.read(data, {type:'array', cellDates: true});
                    const sheet = workbook.Sheets[workbook.SheetNames[0]];
                    const json = XLSX.utils.sheet_to_json(sheet);
                    document.getElementById('rows-box').innerHTML = '';
                    rows = [];
                    json.forEach(row => {
                        addRow({
                            n: row[KEYS.n] || '',
                            loc: normalizeExecutionPlace(row[KEYS.loc]),
                            s: parseExcelDate(row[KEYS.s]),
                            e: parseExcelDate(row[KEYS.e]),
                            st: normalizeStatus(row[KEYS.st]),
                            sp: row[KEYS.sp] || DDL.sup[0],
                            r: row[KEYS.r] || DDL.room[0],
                            f: normalizeFloor(row[KEYS.f]),
                            p: normalizePeriod(row[KEYS.p])
                        });
                    });
                    updateIdx();
                    updateWeekNumberFromRows();
                    invalidateGeneratedFiles();
                    showToast(`تم استيراد ${rows.length} أنشطة تدريبية`, 'success');
                    e.target.value = '';
                } catch(err) { showToast("خطأ في قراءة الملف: " + err.message, 'error'); }
            };
            rd.readAsArrayBuffer(f);
        }

        // --- 4. Archive ---
        function autoSave() { try { const week = document.getElementById('w-id').value; const owner = getSelectedOwner(); const archive = JSON.parse(localStorage.getItem(ARCH_KEY) || '{}'); archive[week] = { rows: rows, owner: owner }; localStorage.setItem(ARCH_KEY, JSON.stringify(archive)); } catch(e) {} }
        function loadFromArchive(week) { const archive = JSON.parse(localStorage.getItem(ARCH_KEY) || '{}'); if(archive[week]) { const data = archive[week]; rows = []; document.getElementById('rows-box').innerHTML = ''; (data.rows || data).forEach(r => addRow(r)); document.getElementById('w-id').value = week; document.getElementById('owner-select').value = data.owner || ""; updateIdx(); invalidateGeneratedFiles(); showPage('generator'); showToast(`تم تحميل الأسبوع ${week}`, 'info'); } }
        function deleteFromArchive(week) { if(confirm(`حذف أرشيف الأسبوع ${week}؟`)) { const archive = JSON.parse(localStorage.getItem(ARCH_KEY) || '{}'); delete archive[week]; localStorage.setItem(ARCH_KEY, JSON.stringify(archive)); refreshArchiveView(); initWeekNumber(); showToast("تم الحذف", 'info'); } }
        function refreshArchiveView() { const container = document.getElementById('archive-list'); const noData = document.getElementById('no-archive'); const archive = JSON.parse(localStorage.getItem(ARCH_KEY) || '{}'); const keys = Object.keys(archive); container.innerHTML = ''; if(keys.length === 0) { noData.classList.remove('hidden'); return; } noData.classList.add('hidden'); keys.sort((a,b) => b - a).forEach(key => { const entry = archive[key]; const owner = entry.owner || 'غير محدد'; const count = (entry.rows || entry).length; const div = document.createElement('div'); div.className = "bg-slate-50 p-4 rounded-lg border hover:shadow-md transition-shadow"; div.innerHTML = `<div class="flex justify-between items-center mb-2"><h4 class="font-bold text-[#2c6060] text-lg">الأسبوع ${key}</h4><span class="text-xs bg-slate-200 px-2 py-1 rounded">${count} دورات</span></div><div class="text-xs text-slate-500 mb-3">بواسطة: <span class="font-bold text-slate-700">${owner}</span></div><div class="flex gap-2 mt-4"><button onclick="loadFromArchive('${key}')" class="flex-1 btn-main text-xs">📝 تعديل</button><button onclick="deleteFromArchive('${key}')" class="btn-outline text-xs text-red-500 border-red-500 hover:bg-red-50">🗑️ حذف</button></div>`; container.appendChild(div); }); }

        // --- 5. Save Logic ---
        
        async function generateScreenZipBlob(valid, week) {
            const stats = getStats(valid);
            const zip = new JSZip();
            const chunksS = chunk(valid, 4);
            for(let i=0; i<chunksS.length; i++) {
                const canvas = document.createElement('canvas'); canvas.width = 2160; canvas.height = 3840;
                const ctx = canvas.getContext('2d'); drawScreenCard(ctx, chunksS[i], stats, i+1, chunksS.length, i*4, logoImage);
                const blob = await new Promise(resolve => canvas.toBlob(resolve, 'image/jpeg', 0.95));
                zip.file('Week-' + week + '-Screen-' + (i+1) + '.jpg', blob);
            }
            return zip.generateAsync({type: "blob"});
        }

        async function startScreenSave() {
            if (!validateOwnerSelection()) return;
            const valid = rows.filter(x => x.n.trim() !== ""); if(valid.length === 0) { showToast("لا توجد بيانات!", 'error'); return; }
            toggleLoader(true, "جاري التوليد...");
            try {
                if(document.fonts && document.fonts.ready) await document.fonts.ready;
                updateWeekNumberFromRows();
                autoSave();
                const week = document.getElementById('w-id').value;
                const content = await generateScreenZipBlob(valid, week);
                lastZipBlob = content;
                lastScreenSignature = getScreenGenerationSignature();
                updateShareButtons();
                saveAs(content, 'NFDP_Week_' + week + '_Screens.zip'); showToast("تم الحفظ!", 'success');
            } catch (err) { console.error(err); showToast("خطأ: " + err.message, 'error'); } finally { toggleLoader(false); }
        }

        async function generateAgentReportBlob(valid) {
            const stats = getStats(valid);
            const canvas = document.createElement('canvas');
            canvas.width = 1080;
            canvas.height = 1920;
            const ctx = canvas.getContext('2d');
            drawAgentReport(ctx, valid, stats, logoImage);
            return new Promise(resolve => canvas.toBlob(resolve, 'image/jpeg', 0.9));
        }

        async function startAgentReport() {
            if (!validateOwnerSelection()) return;
            const valid = rows.filter(x => x.n.trim() !== ""); if(valid.length === 0) { showToast("لا توجد بيانات!", 'error'); return; }
            toggleLoader(true, "جاري التوليد...");
            try {
                if(document.fonts && document.fonts.ready) await document.fonts.ready;
                updateWeekNumberFromRows();
                autoSave();
                const week = document.getElementById('w-id').value;
                const blob = await generateAgentReportBlob(valid);
                lastBlob = blob;
                lastReportSignature = getAgentReportGenerationSignature();
                updateShareButtons();
                saveAs(blob, `NFDP_Week_${week}_Agent_Report.jpg`);
                showToast("تم التوليد!", 'success');
            } catch (err) { console.error(err); showToast("خطأ: " + err.message, 'error'); } finally { toggleLoader(false); }
        }

        // --- Share Logic ---
        function checkShareSupport() { 
            updateShareButtons();
        }

        function splitBase64Lines(base64) {
            return base64.match(/.{1,76}/g).join('\r\n');
        }

        async function blobToBase64(blob) {
            const buffer = await blob.arrayBuffer();
            let binary = '';
            const bytes = new Uint8Array(buffer);
            const chunkSize = 0x8000;
            for (let i = 0; i < bytes.length; i += chunkSize) {
                binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunkSize));
            }
            return btoa(binary);
        }

        function encodeMailHeader(text) {
            return '=?UTF-8?B?' + btoa(unescape(encodeURIComponent(text))) + '?=';
        }

        async function createScreenEml(zipBlob, week) {
            const boundary = '----NAUSS-SCHEDULE-' + Date.now();
            const recipients = SCREEN_EMAIL_RECIPIENTS.join(', ');
            const subject = 'تحديث صور جدول الدورات التدريبية الأسبوعي';
            const body = '<div dir="rtl" style="font-family:Arial,sans-serif;line-height:1.8">' +
                '<p>الزملاء الكرام،</p>' +
                '<p>السلام عليكم ورحمة الله وبركاته،</p>' +
                '<p>مرفق لكم صور جدول الدورات التدريبية للأسبوع ' + week + '، نأمل التكرم برفعها على شاشة مدخل المبنى، واستبدال النسخة السابقة بها، بما يضمن عرض الجدول المحدث للبرامج التدريبية المعتمدة.</p>' +
                '<div dir="ltr" style="text-align:left;margin-top:18px">' +
                '<p>Dear Colleagues,</p>' +
                '<p>Attached are the updated images for week ' + week + ' training courses schedule. Kindly upload them to the building entrance display screen and replace the previous version to ensure that the latest approved training schedule is shown.</p>' +
                '</div>' +
                '<p>مع خالص الشكر والتقدير،</p>' +
                '<p>فريق إدارة عمليات التدريب</p>' +
                '</div>';
            const zipBase64 = splitBase64Lines(await blobToBase64(zipBlob));
            const fileName = 'NFDP_Week_' + week + '_Screens.zip';
            const eml = [
                'To: ' + recipients,
                'Subject: ' + encodeMailHeader(subject),
                'MIME-Version: 1.0',
                'Content-Type: multipart/mixed; boundary="' + boundary + '"',
                '',
                '--' + boundary,
                'Content-Type: text/html; charset=UTF-8',
                'Content-Transfer-Encoding: 8bit',
                '',
                body,
                '',
                '--' + boundary,
                'Content-Type: application/zip; name="' + fileName + '"',
                'Content-Transfer-Encoding: base64',
                'Content-Disposition: attachment; filename="' + fileName + '"',
                '',
                zipBase64,
                '',
                '--' + boundary + '--'
            ].join('\r\n');
            return new Blob([eml], { type: 'message/rfc822;charset=utf-8' });
        }



        async function createAgentReportEml(reportBlob, week) {
            const boundary = '----NAUSS-AGENT-REPORT-' + Date.now();
            const subject = 'جدول الدورات التدريبية للأسبوع القادم';
            const body = '<div dir="rtl" style="font-family:Arial,sans-serif;line-height:1.8">' +
                '<p>سعادة وكيل الجامعة للتدريب سلّمه الله،</p>' +
                '<p>السلام عليكم ورحمة الله وبركاته،</p>' +
                '<p>نرفق لسعادتكم تقرير جدول الدورات التدريبية للأسبوع ' + week + '، متضمنًا البرامج المجدولة، ومواقع التنفيذ، والحالة التشغيلية، ومنسقي التدريب، وذلك للاطلاع والتوجيه بما ترونه مناسبًا.</p>' +
                '<div dir="ltr" style="text-align:left;margin-top:18px">' +
                '<p>Your Excellency,</p>' +
                '<p>Attached is the training courses schedule report for week ' + week + ', including the scheduled programmes, delivery locations, operational status, and assigned training coordinators, for your kind review and guidance as deemed appropriate.</p>' +
                '</div>' +
                '<p>وتفضلوا سعادتكم بقبول خالص الاحترام والتقدير،</p>' +
                '<p>فريق إدارة عمليات التدريب<br>وكالة الجامعة للتدريب</p>' +
                '</div>';
            const reportBase64 = splitBase64Lines(await blobToBase64(reportBlob));
            const fileName = 'NFDP_Week_' + week + '_Agent_Report.jpg';
            const eml = [
                'To: ' + AGENT_REPORT_RECIPIENTS.join(', '),
                'Subject: ' + encodeMailHeader(subject),
                'MIME-Version: 1.0',
                'Content-Type: multipart/mixed; boundary="' + boundary + '"',
                '',
                '--' + boundary,
                'Content-Type: text/html; charset=UTF-8',
                'Content-Transfer-Encoding: 8bit',
                '',
                body,
                '',
                '--' + boundary,
                'Content-Type: image/jpeg; name="' + fileName + '"',
                'Content-Transfer-Encoding: base64',
                'Content-Disposition: attachment; filename="' + fileName + '"',
                '',
                reportBase64,
                '',
                '--' + boundary + '--'
            ].join('\r\n');
            return new Blob([eml], { type: 'message/rfc822;charset=utf-8' });
        }

        async function shareScreenImages() {
            if (!validateOwnerSelection()) return;
            if (!isScreenShareReady()) {
                showToast('لا يمكن مشاركة الملف إلا بعد توليد الملف بشكله النهائي', 'error');
                return;
            }
            toggleLoader(true, "جاري تجهيز مسودة البريد...");
            try {
                const week = document.getElementById('w-id').value;
                const emlBlob = await createScreenEml(lastZipBlob, week);
                saveAs(emlBlob, 'Screen_Schedule_Week_' + week + '.eml');
                showToast("تم تنزيل مسودة البريد", 'success');
            } catch (error) { console.error(error); showToast("خطأ في إنشاء مسودة البريد", 'error'); }
            finally { toggleLoader(false); }
        }

        async function shareAgentReport() {
            if (!validateOwnerSelection()) return;
            if (!isAgentReportShareReady()) {
                showToast('لا يمكن مشاركة الملف إلا بعد توليد الملف بشكله النهائي', 'error');
                return;
            }
            toggleLoader(true, "جاري تجهيز مسودة بريد الوكيل...");
            try {
                const week = document.getElementById('w-id').value;
                const emlBlob = await createAgentReportEml(lastBlob, week);
                saveAs(emlBlob, 'Agent_Report_Week_' + week + '.eml');
                showToast("تم تنزيل مسودة بريد الوكيل", 'success');
            } catch (error) { console.error(error); showToast("خطأ في إنشاء مسودة بريد الوكيل", 'error'); }
            finally { toggleLoader(false); }
        }

        // --- 6. Canvas Drawing ---
        
        function wrapTextSmart(ctx, text, maxWidth, maxLines, baseFontSize) {
            let fontSize = baseFontSize;
            const len = text.length;
            if (len > 90) fontSize = baseFontSize * 0.6; else if (len > 60) fontSize = baseFontSize * 0.8;
            ctx.font = `bold ${fontSize}px Cairo`;
            let lines = []; let words = text.split(' '); let currentLine = '';
            for (let i = 0; i < words.length; i++) {
                let word = words[i]; let testLine = currentLine + (currentLine ? ' ' : '') + word;
                if (ctx.measureText(testLine).width > maxWidth && currentLine) { lines.push(currentLine); currentLine = word; if (lines.length >= maxLines) break; } else { currentLine = testLine; }
            }
            if (lines.length < maxLines) lines.push(currentLine);
            if (lines.length > maxLines) lines = lines.slice(0, maxLines);
            let last = lines[lines.length - 1];
            while (ctx.measureText(last + '...').width > maxWidth && last.length > 0) last = last.substring(0, last.length - 1);
            lines[lines.length - 1] = last + (len > 80 ? '...' : '');
            const lineHeight = fontSize * 1.4; 
            return { lines: lines, fontSize: fontSize, lineHeight: lineHeight };
        }

        function drawRoundedRect(ctx, x, y, width, height, radius) {
            ctx.beginPath(); ctx.moveTo(x + radius, y); ctx.lineTo(x + width - radius, y); ctx.quadraticCurveTo(x + width, y, x + width, y + radius); ctx.lineTo(x + width, y + height - radius); ctx.quadraticCurveTo(x + width, y + height, x + width - radius, y + height); ctx.lineTo(x + radius, y + height); ctx.quadraticCurveTo(x, y + height, x, y + height - radius); ctx.lineTo(x, y + radius); ctx.quadraticCurveTo(x, y, x + radius, y); ctx.closePath(); ctx.fill();
        }
        
        // ---------------------------------------------------------
        // Agent Report Drawing (Cleaned)
        // ---------------------------------------------------------
        function drawAgentReport(ctx, list, stats, logo) {
            const W = 1080, H = 1920;

            // 1. BACKGROUND
            ctx.fillStyle = UNI.bodyBg; 
            ctx.fillRect(0, 0, W, H);

            // 2. PATTERN
            ctx.save();
            ctx.fillStyle = "rgba(255, 255, 255, 0.2)"; 
            const step = 60;
            for (let i = -step; i < W + step; i += step) {
                for (let j = -step; j < H + step; j += step) {
                    const offsetX = (j / step) % 2 === 0 ? 0 : step / 2;
                    ctx.beginPath();
                    ctx.arc(i + offsetX, j, 15, 0, Math.PI * 2);
                    ctx.fill();
                }
            }
            ctx.restore();

            // 3. HEADER BLOCK
            let y = 50;
            const logoMaxW = 450;
            let logoW = logoMaxW;
            let logoH = 0;
            if (logo && logo.width > 0) { logoH = logoW * (logo.height / logo.width); }
            const headerHeight = 50 + logoH + 30 + 50 + 30 + 90 + 40;

            ctx.fillStyle = UNI.green;
            ctx.fillRect(0, 0, W, headerHeight);

            // 4. LOGO
            if (logo && logo.width > 0) { 
                const xPos = (W - logoW) / 2;
                ctx.drawImage(logo, xPos, y, logoW, logoH); 
                y += logoH + 30; 
            } else { y += 100; }
            
            // 5. TITLE
            ctx.textAlign = 'center';
            ctx.fillStyle = UNI.gold; 
            ctx.font = 'bold 45px Cairo';
            ctx.fillText("جدول الدورات التدريبية", W/2, y);
            y += 80;

            // 6. STATS
            const boxW = 220; const boxH = 90; const gap = 20;
            const startX = (W - ((boxW * 4) + (gap * 3))) / 2;
            
            const drawStatBox = (x, label, val) => {
                ctx.fillStyle = 'rgba(255, 255, 255, 0.2)'; 
                drawRoundedRect(ctx, x, y, boxW, boxH, 15);
                ctx.strokeStyle = 'rgba(255, 255, 255, 0.4)'; ctx.lineWidth = 2;
                ctx.strokeRect(x, y, boxW, boxH);
                ctx.fillStyle = '#ffffff';
                ctx.textAlign = 'center';
                ctx.font = 'bold 40px Cairo';
                ctx.fillText(val, x + boxW/2, y + 45);
                ctx.font = 'normal 18px Cairo';
                ctx.fillText(label, x + boxW/2, y + 70);
            };

            let bX = startX;
            drawStatBox(bX, 'إجمالي', stats.t); bX += boxW + gap;
            drawStatBox(bX, 'جديدة', stats.n); bX += boxW + gap;
            drawStatBox(bX, 'مستمرة', stats.c); bX += boxW + gap;
            drawStatBox(bX, 'خارجية', stats.i);

            y += boxH;

            // END OF GREEN HEADER.
            // Gap of 2 lines
            y = headerHeight + 70;

            // 7. INTRO TEXT (Updated Script - No "البيان")
            ctx.textAlign = 'right';
            ctx.fillStyle = '#0f172a'; 
            ctx.font = 'normal 28px Cairo';
            
            const introLines = [
                "سعادة وكيل التدريب سلمه الله ..",
                "السلام عليكم ورحمة الله وبركاته",
                "",
                "يسرنا أن نضع بين يديكم جدول الدورات التدريبية للاسبوع القادم وهو على النحو التالي :"
            ];
            introLines.forEach(line => {
                ctx.fillText(line, W - 40, y);
                y += 45;
            });
            y += 20; // Small gap before list

            // 8. COURSES LIST
            const cardW = W - 80; const cardX = 40; const cardH = 115;
            
            for (let r = 0; r < list.length; r++) {
                const item = list[r];
                if (y + cardH > H - 150) break;

                ctx.fillStyle = '#ffffff';
                ctx.shadowColor = 'rgba(0,0,0,0.08)'; ctx.shadowBlur = 10; ctx.shadowOffsetY = 5;
                drawRoundedRect(ctx, cardX, y, cardW, cardH, 15);
                ctx.shadowColor = 'transparent';

                // Number
                ctx.fillStyle = UNI.green; 
                ctx.beginPath(); ctx.arc(W - 75, y + 35, 20, 0, 2*Math.PI); ctx.fill();
                ctx.fillStyle = '#fff'; ctx.font = 'bold 18px Cairo'; ctx.textAlign = 'center';
                ctx.fillText(r+1, W - 75, y + 41);

                // Title
                ctx.fillStyle = '#1e293b'; ctx.font = 'bold 22px Cairo'; ctx.textAlign = 'right';
                let title = item.n; while(ctx.measureText(title).width > cardW - 100 && title.length > 0) title = title.substring(0, title.length-1);
                if (title !== item.n) title += "...";
                ctx.fillText(title, W - 100, y + 32);

                // Line 2: Details Right, Date Left
                ctx.font = 'normal 16px Cairo';
                ctx.fillStyle = '#64748b';
                const details = `${item.st}  •  ${item.p}  •  ${item.loc}`;
                ctx.textAlign = 'right';
                ctx.fillText(details, W - 100, y + 58);

                // Date on Left
                const sDate = formatDateShort(item.s);
                const eDate = formatDateShort(item.e);
                const dateStr = (sDate && eDate) ? `${sDate} - ${eDate}` : (sDate || eDate);
                ctx.fillStyle = UNI.green;
                ctx.textAlign = 'left';
                ctx.fillText(dateStr, 100, y + 58);

                // Location and Coordinator
                ctx.fillStyle = '#475569'; ctx.font = 'normal 15px Cairo';
                ctx.textAlign = 'right';
                ctx.fillText(`${item.f}  •  ${item.r}`, W - 100, y + 85);
                ctx.fillStyle = UNI.supRed; ctx.font = 'bold 15px Cairo';
                ctx.textAlign = 'left';
                ctx.fillText(`👤 ${item.sp}`, 100, y + 85);
                
                y += cardH + 15;
            }

            // 9. FOOTER
            ctx.fillStyle = '#64748b'; ctx.textAlign = 'center';
            ctx.font = 'normal 26px Cairo';
            ctx.fillText("وتفضلوا بقبول فائق الاحترام والتقدير", W/2, H - 70);
            ctx.fillStyle = UNI.green;
            ctx.font = 'bold 26px Cairo';
            ctx.fillText("فريق إدارة عمليات التدريب", W/2, H - 30);
        }

        // ---------------------------------------------------------
        // Screen Card Drawing (Templates)
        // ---------------------------------------------------------
        function drawScreenCard(ctx, list, stats, pn, tp, si, logo) {
            const W = 2160, H = 3840;
            drawBackground(ctx, W, H);

            const headerPad = 60;
            const activeLogo = (currentTemplate === 3 && darkLogoImage) ? darkLogoImage : logo;
            const logoTargetW = 1000; 
            let logoH = 0;
            if (activeLogo && activeLogo.width > 0) { logoH = logoTargetW * (activeLogo.height / activeLogo.width); }
            const titleFontSize = 100;
            const titleH = titleFontSize + 20;
            const totalHeaderH = headerPad + logoH + (logoH > 0 ? 40 : 0) + titleH + headerPad;

            if (currentTemplate !== 6) {
                 if(currentTemplate !== 4 && currentTemplate !== 5) {
                    if(currentTemplate === 2) ctx.fillStyle = UNI.green;
                    else ctx.fillStyle = currentTemplate === 3 ? 'transparent' : '#2c6060';
                    if(currentTemplate === 3) { let g = ctx.createLinearGradient(0,0,W,0); g.addColorStop(0,"#14b8a6"); g.addColorStop(1,"#0ea5e9"); ctx.fillStyle = g; }
                    ctx.fillRect(0, 0, W, totalHeaderH);
                }
                if(currentTemplate === 4) { ctx.fillStyle = UNI.green; ctx.fillRect(0, 0, W, totalHeaderH); ctx.fillStyle = UNI.gold; ctx.fillRect(0, totalHeaderH - 15, W, 15); }
            }

            let currentDrawY = headerPad;
            if (activeLogo && activeLogo.width > 0) { const logoX = (W - logoTargetW) / 2; ctx.drawImage(activeLogo, logoX, currentDrawY, logoTargetW, logoH); currentDrawY += logoH + 40; }
            
            if(currentTemplate === 6) ctx.fillStyle = '#ffffff';
            else if(currentTemplate === 2 || currentTemplate === 4) ctx.fillStyle = UNI.gold;
            else if(currentTemplate === 5) ctx.fillStyle = '#3e2723';
            else ctx.fillStyle = '#ffffff';
            
            ctx.textAlign = 'center';
            ctx.font = `bold ${titleFontSize}px Cairo`;
            ctx.fillText("جدول الدورات التدريبية", W/2, currentDrawY + titleFontSize);
            const headerEndY = totalHeaderH;

            drawStatsBlock(ctx, stats, W, headerEndY + 50);

            if(currentTemplate === 6) ctx.fillStyle = 'rgba(255,255,255,0.9)';
            else if(currentTemplate === 2 || currentTemplate === 4) ctx.fillStyle = UNI.green;
            else if(currentTemplate === 5) ctx.fillStyle = '#3e2723';
            else ctx.fillStyle = '#475569';
            ctx.font = 'normal 45px Cairo';
            ctx.textAlign = 'center';
            const introText = getScreenIntroText();
            let introY = headerEndY + 400;
            if (introText) {
                const introLines = wrapTextSimple(ctx, introText, W - 200);
                introLines.forEach(line => { ctx.fillText(line, W/2, introY); introY += 60; });
                introY += 80;
            }

            let y = introText ? introY : headerEndY + 360;
            const cardW = 2000; const startX = 80; const cardGap = 40;
            const remainingH = H - y - 150; 
            const cardH = (remainingH - (cardGap * 3)) / 4;

            for(let i=0; i<list.length; i++) { const c = list[i]; if(c) { drawCourseCard(ctx, c, startX, y, cardW, cardH, si+i+1, 'screen', i); y += cardH + cardGap; } }
            drawFooter(ctx, W, H);
        }

        function wrapTextSimple(ctx, text, maxWidth) {
            const words = text.split(' '); const lines = []; let currentLine = '';
            words.forEach(word => { const testLine = currentLine + (currentLine ? ' ' : '') + word; if (ctx.measureText(testLine).width > maxWidth && currentLine) { lines.push(currentLine); currentLine = word; } else { currentLine = testLine; } });
            lines.push(currentLine); return lines;
        }

        function drawBackground(ctx, W, H) {
            if (currentTemplate === 6) {
                const grd = ctx.createLinearGradient(0, 0, W, H); grd.addColorStop(0, UNI.green); grd.addColorStop(1, UNI.greenMid); ctx.fillStyle = grd; ctx.fillRect(0, 0, W, H);
                ctx.fillStyle = 'rgba(255,255,255,0.05)'; ctx.beginPath(); ctx.arc(W * 0.8, H * 0.2, 600, 0, Math.PI * 2); ctx.fill();
            } else if (currentTemplate === 4) {
                ctx.fillStyle = '#f5f5f7'; ctx.fillRect(0, 0, W, H);
            } else if (currentTemplate === 5) {
                ctx.fillStyle = '#f4e4bc'; ctx.fillRect(0, 0, W, H);
                ctx.strokeStyle = UNI.paperPattern; ctx.lineWidth = 2;
                const step = 100;
                for(let i = 0; i < W; i+=step) { for(let j = 0; j < H; j+=step) { ctx.strokeRect(i, j, 50, 50); ctx.beginPath(); ctx.arc(i+25, j+25, 10, 0, Math.PI*2); ctx.stroke(); } }
                ctx.strokeStyle = "#a1887f"; ctx.lineWidth = 15; ctx.strokeRect(30, 30, W-60, H-60);
            } else {
                ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, W, H);
                if(currentTemplate === 2) { ctx.fillStyle = UNI.white; ctx.fillRect(0, 0, W, H); ctx.strokeStyle = UNI.gold; ctx.lineWidth = 0.5; ctx.globalAlpha = 0.2; for(let i=0; i<W; i+=40) { ctx.beginPath(); ctx.moveTo(i, 0); ctx.lineTo(i, H); ctx.stroke(); } ctx.globalAlpha = 1.0; }
                if(currentTemplate === 3) { ctx.fillStyle = '#f0fdfa'; ctx.fillRect(0, 0, W, H); }
            }
        }

        function drawStatsBlock(ctx, stats, W, Y) {
            const isExec = W < 2000; const sc = isExec ? 0.5 : 1;
            const boxW = 450*sc, boxH = 200*sc, gap = 35*sc; const totalW = (boxW * 4) + (gap * 3); let startX = (W - totalW) / 2;
            
            const drawBox = (x, bg, txt, val, txtColor = '#fff') => {
                if (currentTemplate === 6) { ctx.fillStyle = 'rgba(255, 255, 255, 0.2)'; drawRoundedRect(ctx, x, Y, boxW, boxH, 15*sc); ctx.strokeStyle = 'rgba(255, 255, 255, 0.5)'; ctx.lineWidth = 4; ctx.strokeRect(x, Y, boxW, boxH); txtColor = '#ffffff'; }
                else if (currentTemplate === 4) { ctx.fillStyle = '#ffffff'; drawRoundedRect(ctx, x, Y, boxW, boxH, 20); ctx.shadowColor = 'rgba(0,0,0,0.1)'; ctx.shadowBlur = 20; ctx.shadowOffsetY = 5; txtColor = '#1d1d1f'; }
                else if(currentTemplate === 5) {
                    ctx.fillStyle = '#fffbf0'; drawRoundedRect(ctx, x, Y, boxW, boxH, 5);
                    ctx.strokeStyle = '#a1887f'; ctx.lineWidth = 2; ctx.strokeRect(x, Y, boxW, boxH);
                    txtColor = '#5d4037';
                }
                else if(currentTemplate === 2 || currentTemplate === 4) { ctx.fillStyle = (currentTemplate === 4 ? UNI.grayLight : UNI.green); drawRoundedRect(ctx, x, Y, boxW, boxH, 15*sc); ctx.strokeStyle = UNI.gold; ctx.lineWidth = 4; ctx.strokeRect(x, Y, boxW, boxH); txtColor = UNI.gold; }
                else { ctx.fillStyle = bg; drawRoundedRect(ctx, x, Y, boxW, boxH, 15*sc); }
                ctx.fillStyle = txtColor; ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
                ctx.font = `bold ${90*sc}px Cairo`; ctx.fillText(val, x + boxW/2, Y + boxH/2 - 25*sc);
                ctx.font = `normal ${40*sc}px Cairo`; ctx.fillText(txt, x + boxW/2, Y + boxH/2 + 35*sc);
            };

            let box3Color = '#d7be8c'; let box3TextColor = '#2c6060';
            if (currentTemplate === 6) { box3Color = 'rgba(255,255,255,0.2)'; box3TextColor = '#fff'; }
            if(currentTemplate === 4) { box3Color = '#f5f5f7'; box3TextColor = '#1d1d1f'; }
            if(currentTemplate === 3) { box3Color = '#0ea5e9'; box3TextColor = '#fff'; }
            if(currentTemplate === 2 || currentTemplate === 4) { box3Color = UNI.grayLight; box3TextColor = UNI.gold; }

            drawBox(startX, '#2c6060', 'إجمالي', stats.t);
            drawBox(startX + boxW + gap, '#c2b59d', 'جديدة', stats.n);
            drawBox(startX + (boxW + gap)*2, box3Color, 'مستمرة', stats.c, box3TextColor);
            
            if (currentTemplate === 6) { ctx.fillStyle = 'rgba(255, 255, 255, 0.2)'; drawRoundedRect(ctx, startX + (boxW + gap)*3, Y, boxW, boxH, 15*sc); ctx.strokeStyle = 'rgba(255, 255, 255, 0.5)'; ctx.lineWidth = 4; ctx.strokeRect(startX + (boxW + gap)*3, Y, boxW, boxH); ctx.fillStyle = '#ffffff'; }
            else if (currentTemplate === 4) { ctx.fillStyle = '#ffffff'; drawRoundedRect(ctx, startX + (boxW + gap)*3, Y, boxW, boxH, 20); ctx.shadowColor='transparent'; ctx.fillStyle = '#1d1d1f'; }
            else if(currentTemplate === 2 || currentTemplate === 4) { ctx.strokeStyle = UNI.gold; ctx.lineWidth = 4; ctx.strokeRect(startX + (boxW + gap)*3, Y, boxW, boxH); ctx.fillStyle = UNI.gold; }
            else if (currentTemplate === 5) { ctx.fillStyle = '#fffbf0'; drawRoundedRect(ctx, startX + (boxW + gap)*3, Y, boxW, boxH, 5); ctx.strokeStyle = '#a1887f'; ctx.lineWidth = 2; ctx.strokeRect(startX + (boxW + gap)*3, Y, boxW, boxH); ctx.fillStyle = '#5d4037'; }
            else { ctx.strokeStyle = '#2c6060'; ctx.lineWidth = 4*sc; ctx.strokeRect(startX + (boxW + gap)*3, Y, boxW, boxH); ctx.fillStyle = '#2c6060'; }
            ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
            ctx.font = `bold ${90*sc}px Cairo`; ctx.fillText(stats.i, startX + (boxW + gap)*3 + boxW/2, Y + boxH/2 - 25*sc);
            ctx.font = `normal ${40*sc}px Cairo`; ctx.fillText('خارجية', startX + (boxW + gap)*3 + boxW/2, Y + boxH/2 + 35*sc);
        }

        function drawFooter(ctx, W, H) {
            const footerH = 120; let bgColor = '#e2e8f0'; let textColor = '#64748b';
            if (currentTemplate === 6) { bgColor = 'rgba(0,0,0,0.3)'; textColor = '#ffffff'; }
            if (currentTemplate === 4) { bgColor = '#f5f5f7'; textColor = '#8e8e93'; }
            if(currentTemplate === 2 || currentTemplate === 4) { bgColor = UNI.green; textColor = UNI.gold; }
            if(currentTemplate === 5) { bgColor = '#a1887f'; textColor = '#3e2723'; }
            ctx.fillStyle = bgColor; ctx.fillRect(0, H - footerH, W, footerH);
            ctx.fillStyle = textColor; ctx.textAlign = 'center';
            ctx.font = 'normal 35px Cairo';
            ctx.fillText("وكالة التدريب بجامعة نايف العربية للعلوم الأمنية", W/2, H - 40);
        }

        function drawCourseCard(ctx, c, x, y, w, h, idx, type, loopIndex) {
            const isScreen = type === 'screen'; const sc = isScreen ? 1 : 0.5;
            const padding = 50*sc; const radius = (currentTemplate === 4 ? 40 : 20) * sc;
            
            let cardBg = '#ffffff'; let cardText = '#1e293b'; let cardShadow = 'rgba(0,0,0,0.1)'; let cardBorder = '#e2e8f0';

            if (currentTemplate === 6) {
                ctx.shadowColor = 'rgba(0,0,0,0.2)'; ctx.shadowBlur = 30; ctx.shadowOffsetY = 10;
                ctx.fillStyle = 'rgba(255, 255, 255, 0.85)'; drawRoundedRect(ctx, x, y, w, h, radius);
                ctx.shadowColor = 'transparent';
                ctx.strokeStyle = 'rgba(255, 255, 255, 0.6)'; ctx.lineWidth = 4;
                ctx.beginPath(); ctx.moveTo(x + radius, y); ctx.arcTo(x + w, y, x + w, y + h, radius); ctx.arcTo(x + w, y + h, x, y + h, radius); ctx.arcTo(x, y + h, x, y, radius); ctx.arcTo(x, y, x + w, y, radius); ctx.stroke();
                cardText = '#016564';
            } else if (currentTemplate === 4) {
                ctx.shadowColor = 'rgba(0,0,0,0.15)'; ctx.shadowBlur = 40; ctx.shadowOffsetY = 10;
                ctx.fillStyle = '#ffffff'; drawRoundedRect(ctx, x, y, w, h, radius);
                ctx.shadowColor = 'transparent';
                cardText = '#1d1d1f';
            } else if (currentTemplate === 5) {
                ctx.fillStyle = '#fffef5'; drawRoundedRect(ctx, x, y, w, h, 5);
                ctx.strokeStyle = '#c4a77d'; ctx.lineWidth = 2; ctx.strokeRect(x, y, w, h);
                cardText = '#3e2723';
            } else if(currentTemplate === 2) { cardBg = '#ffffff'; cardBorder = UNI.gold; cardText = UNI.green; cardShadow = `rgba(208, 178, 132, 0.3)`; }
            else if(currentTemplate === 4) {
                ctx.fillStyle = UNI.gold; drawRoundedRect(ctx, x + 15, y + 15, w, h, radius);
                ctx.fillStyle = '#ffffff'; drawRoundedRect(ctx, x, y, w, h, radius);
                ctx.strokeStyle = UNI.green; ctx.lineWidth = 3; ctx.beginPath(); ctx.moveTo(x + radius, y); ctx.arcTo(x + w, y, x + w, y + h, radius); ctx.arcTo(x + w, y + h, x, y + h, radius); ctx.arcTo(x, y + h, x, y, radius); ctx.arcTo(x, y, x + w, y, radius); ctx.stroke();
                cardText = UNI.green;
            }
            
            if(currentTemplate !== 4 && currentTemplate !== 6) {
                ctx.fillStyle = cardBg; ctx.shadowColor = cardShadow; ctx.shadowBlur = 25*sc; ctx.shadowOffsetY = 5*sc; drawRoundedRect(ctx, x, y, w, h, radius); ctx.shadowColor = 'transparent';
                if(currentTemplate !== 3 && currentTemplate !== 5) { ctx.strokeStyle = cardBorder; ctx.lineWidth = currentTemplate === 2 ? 3 : 2*sc; ctx.beginPath(); ctx.moveTo(x + radius, y); ctx.arcTo(x + w, y, x + w, y + h, radius); ctx.arcTo(x + w, y + h, x, y + h, radius); ctx.arcTo(x, y + h, x, y, radius); ctx.arcTo(x, y, x + w, y, radius); ctx.stroke(); }
            }

            const isVibrantScreen = currentTemplate === 3 && isScreen;
            const isIosTemplate = currentTemplate === 4;
            const baseTitleSize = isVibrantScreen ? 70 : (isScreen ? 65 : 40);
            const dateFontSize = isVibrantScreen ? 38 : (isScreen ? 35 : 24);
            const gridH = isIosTemplate ? (isScreen ? 220 : 136) : (isVibrantScreen ? 170 : (isScreen ? 150 : 100));
            const supFontSize = isIosTemplate ? (isScreen ? 39 : 26) : (isVibrantScreen ? 44 : (isScreen ? 35 : 24));
            const gapSize = isVibrantScreen ? 42 : (isScreen ? 25 : 15);
            const titleDateGap = isIosTemplate ? (isScreen ? 34 : 20) : gapSize;
            const dateGridGap = isIosTemplate ? (isScreen ? 32 : 18) : gapSize;
            const gridSupGap = isIosTemplate ? (isScreen ? 30 : 18) : gapSize;
            const fullTitle = c.n + (isExternalExecution(c.loc) ? ' 🌍' : '');
            const result = wrapTextSmart(ctx, fullTitle, w - padding*2, 2, baseTitleSize);
            const titleLines = result.lines; const titleFontSize = result.fontSize; const lineHeight = result.lineHeight;
            const titleH = titleLines.length * lineHeight;
            
            const totalContentH = titleH + titleDateGap + dateFontSize + dateGridGap + gridH + gridSupGap + supFontSize;
            let currentY = y + (h - totalContentH) / 2;
            const centerX = x + w / 2;

            ctx.font = `bold ${titleFontSize}px Cairo`; ctx.fillStyle = cardText; ctx.textAlign = 'center';
            let textY = currentY + titleFontSize;
            titleLines.forEach((line, index) => { ctx.fillText(line, centerX, textY); textY += lineHeight; });
            currentY += titleH + titleDateGap;

            if (currentTemplate === 6) ctx.fillStyle = 'rgba(1, 101, 100, 0.8)';
            else if(currentTemplate === 2) ctx.fillStyle = UNI.greenMid;
            else if(currentTemplate === 5) ctx.fillStyle = '#555';
            else ctx.fillStyle = '#64748b';
            const dateText = isVibrantScreen ? `${c.s} إلى ${c.e}` : `${c.s} - ${c.e}`;
            ctx.font = `normal ${dateFontSize}px Cairo`; ctx.fillText(dateText, centerX, currentY + dateFontSize); currentY += dateFontSize + dateGridGap;

            const infoPadding = isVibrantScreen ? 70*sc : padding;
            let gridStartX = x + infoPadding; const boxGap = isVibrantScreen ? 28*sc : 20*sc; const boxW = (w - infoPadding*2 - (boxGap*4)) / 5;
            const drawLineIcon = (iconType, cx, cy, size, color) => {
                ctx.save();
                ctx.strokeStyle = color;
                ctx.lineWidth = Math.max(3, size * 0.12);
                ctx.lineCap = 'round';
                ctx.lineJoin = 'round';
                const s = size;
                switch (iconType) {
                    case 'location': {
                        ctx.beginPath();
                        ctx.arc(cx, cy - s * 0.12, s * 0.18, 0, Math.PI * 2);
                        ctx.stroke();
                        ctx.beginPath();
                        ctx.moveTo(cx, cy + s * 0.42);
                        ctx.lineTo(cx - s * 0.18, cy + s * 0.08);
                        ctx.quadraticCurveTo(cx - s * 0.28, cy - s * 0.1, cx, cy - s * 0.3);
                        ctx.quadraticCurveTo(cx + s * 0.28, cy - s * 0.1, cx + s * 0.18, cy + s * 0.08);
                        ctx.closePath();
                        ctx.stroke();
                        break;
                    }
                    case 'status': {
                        ctx.beginPath();
                        ctx.arc(cx, cy, s * 0.3, 0, Math.PI * 2);
                        ctx.stroke();
                        ctx.beginPath();
                        ctx.moveTo(cx - s * 0.12, cy + s * 0.02);
                        ctx.lineTo(cx - s * 0.02, cy + s * 0.14);
                        ctx.lineTo(cx + s * 0.16, cy - s * 0.1);
                        ctx.stroke();
                        break;
                    }
                    case 'period': {
                        ctx.beginPath();
                        ctx.arc(cx, cy, s * 0.3, 0, Math.PI * 2);
                        ctx.stroke();
                        ctx.beginPath();
                        ctx.moveTo(cx, cy);
                        ctx.lineTo(cx, cy - s * 0.14);
                        ctx.moveTo(cx, cy);
                        ctx.lineTo(cx + s * 0.12, cy + s * 0.08);
                        ctx.stroke();
                        break;
                    }
                    case 'room': {
                        ctx.strokeRect(cx - s * 0.28, cy - s * 0.22, s * 0.56, s * 0.44);
                        ctx.beginPath();
                        ctx.moveTo(cx - s * 0.34, cy - s * 0.22);
                        ctx.lineTo(cx, cy - s * 0.42);
                        ctx.lineTo(cx + s * 0.34, cy - s * 0.22);
                        ctx.stroke();
                        ctx.beginPath();
                        ctx.moveTo(cx, cy + s * 0.22);
                        ctx.lineTo(cx, cy - s * 0.02);
                        ctx.stroke();
                        break;
                    }
                    case 'floor': {
                        ctx.beginPath();
                        ctx.moveTo(cx - s * 0.28, cy + s * 0.18);
                        ctx.lineTo(cx - s * 0.08, cy + s * 0.18);
                        ctx.lineTo(cx - s * 0.08, cy);
                        ctx.lineTo(cx + s * 0.12, cy);
                        ctx.lineTo(cx + s * 0.12, cy - s * 0.18);
                        ctx.lineTo(cx + s * 0.3, cy - s * 0.18);
                        ctx.stroke();
                        break;
                    }
                    default: {
                        ctx.beginPath();
                        ctx.arc(cx, cy, s * 0.08, 0, Math.PI * 2);
                        ctx.stroke();
                    }
                }
                ctx.restore();
            };

            const drawInfoBox = (label, value) => {
                const displayValue = value || '-';
                if (currentTemplate === 6) { ctx.fillStyle = 'rgba(1, 101, 100, 0.1)'; drawRoundedRect(ctx, gridStartX, currentY, boxW, gridH, 8*sc); ctx.fillStyle = UNI.green; }
                else if (currentTemplate === 4) { ctx.fillStyle = '#f5f5f7'; drawRoundedRect(ctx, gridStartX, currentY, boxW, gridH, 15*sc); ctx.fillStyle = '#8e8e93'; }
                else if(currentTemplate === 5) {
                    ctx.fillStyle = 'rgba(161, 136, 127, 0.1)'; drawRoundedRect(ctx, gridStartX, currentY, boxW, gridH, 0);
                    ctx.fillStyle = '#5d4037';
                }
                else if(currentTemplate === 2) { ctx.fillStyle = '#f0fdfa'; drawRoundedRect(ctx, gridStartX, currentY, boxW, gridH, 8*sc); ctx.strokeStyle = UNI.gold; ctx.lineWidth = 1; ctx.strokeRect(gridStartX, currentY, boxW, gridH); ctx.fillStyle = UNI.green; }
                else { ctx.fillStyle = '#f8fafc'; drawRoundedRect(ctx, gridStartX, currentY, boxW, gridH, 10*sc); ctx.fillStyle = '#64748b'; }
                
                if (currentTemplate === 4) { ctx.font = `bold 30px Cairo`; ctx.fillText(label, gridStartX + boxW/2, currentY + gridH * 0.35); }
                else { ctx.font = `normal ${isVibrantScreen ? 25*sc : (isScreen ? 23 : 16) * sc}px Cairo`; ctx.textAlign = 'center'; ctx.textBaseline = 'middle'; ctx.fillText(label, gridStartX + boxW/2, currentY + gridH * 0.32); }
                
                ctx.fillStyle = cardText; ctx.font = `bold ${isVibrantScreen ? 34*sc : (isScreen ? 30 : 20) * sc}px Cairo`;
                let safeValue = String(displayValue);
                while (ctx.measureText(safeValue).width > boxW - 24*sc && safeValue.length > 0) safeValue = safeValue.substring(0, safeValue.length - 1);
                if (safeValue !== String(displayValue)) safeValue += '…';
                ctx.fillText(safeValue, gridStartX + boxW/2, currentY + gridH * 0.7);
                gridStartX += boxW + boxGap;
            };

            const drawIosInfoBoxAt = (boxX, boxY, boxWLocal, boxHLocal, iconType, label, value) => {
                const displayValue = value || '-';
                ctx.fillStyle = '#f5f5f7';
                drawRoundedRect(ctx, boxX, boxY, boxWLocal, boxHLocal, 18*sc);

                const labelFontSize = isScreen ? 19 : 12;
                const valueFontSize = isScreen ? 28 : 18;
                const iconSize = isScreen ? 36 : 22;
                const iconGap = isScreen ? 12 : 8;
                const headerY = boxY + boxHLocal * 0.32;
                const headerColor = '#8e8e93';
                const iconColor = '#C7B08C';

                ctx.textBaseline = 'middle';
                ctx.font = `normal ${labelFontSize}px Cairo`;
                const labelWidth = ctx.measureText(label).width;
                const groupWidth = labelWidth + iconGap + iconSize;
                const groupRightX = boxX + (boxWLocal + groupWidth) / 2;
                const iconCenterX = groupRightX - iconSize / 2;
                const labelCenterX = groupRightX - iconSize - iconGap - labelWidth / 2;

                drawLineIcon(iconType, iconCenterX, headerY, iconSize, iconColor);

                ctx.fillStyle = headerColor;
                ctx.textAlign = 'center';
                ctx.fillText(label, labelCenterX, headerY);

                ctx.fillStyle = cardText;
                ctx.font = `bold ${valueFontSize}px Cairo`;
                let safeValue = String(displayValue);
                while (ctx.measureText(safeValue).width > boxWLocal - 36*sc && safeValue.length > 0) safeValue = safeValue.substring(0, safeValue.length - 1);
                if (safeValue !== String(displayValue)) safeValue += '…';
                ctx.fillText(safeValue, boxX + boxWLocal / 2, boxY + boxHLocal * 0.74);
            };
            
            if (isIosTemplate) {
                const contentW = w - infoPadding * 2;
                const topBoxGap = 22 * sc;
                const bottomBoxGap = 22 * sc;
                const rowBoxH = isScreen ? 96 : 58;
                const rowGap = isScreen ? 24 : 14;
                const topBoxW = (contentW - topBoxGap * 2) / 3;
                const bottomBoxW = (contentW - bottomBoxGap) / 2;
                const topRow = [
                    { icon: 'location', label: 'مكان التنفيذ', value: c.loc },
                    { icon: 'status', label: 'الحالة', value: c.st },
                    { icon: 'period', label: 'الفترة', value: c.p }
                ];
                const bottomRow = [
                    { icon: 'room', label: 'القاعة', value: c.r },
                    { icon: 'floor', label: 'الدور', value: shouldHideFloorValue(c.f) ? '' : c.f }
                ];
                topRow.forEach((item, index) => {
                    const boxX = x + infoPadding + contentW - topBoxW - index * (topBoxW + topBoxGap);
                    drawIosInfoBoxAt(boxX, currentY, topBoxW, rowBoxH, item.icon, item.label, item.value);
                });
                const secondRowY = currentY + rowBoxH + rowGap;
                bottomRow.forEach((item, index) => {
                    const boxX = x + infoPadding + contentW - bottomBoxW - index * (bottomBoxW + bottomBoxGap);
                    drawIosInfoBoxAt(boxX, secondRowY, bottomBoxW, rowBoxH, item.icon, item.label, item.value);
                });
            } else if (currentTemplate === 4) { drawInfoBox('⚡', c.st); drawInfoBox('⏰', c.p); drawInfoBox('🌍', c.loc); if (!shouldHideFloorValue(c.f)) drawInfoBox('🧱', c.f); drawInfoBox('🏛️', c.r); }
            else { drawInfoBox('الحالة', c.st); drawInfoBox('الفترة', c.p); drawInfoBox('مكان التنفيذ', c.loc); if (!shouldHideFloorValue(c.f)) drawInfoBox('الطابق', c.f); drawInfoBox('القاعة', c.r); }
            
            currentY += gridH + gridSupGap;

            ctx.fillStyle = UNI.supRed;
            ctx.font = `${isVibrantScreen || isIosTemplate ? 'bold' : 'normal'} ${supFontSize}px Cairo`; ctx.fillText(`اسم منسق التدريب: ${c.sp}`, centerX, currentY + supFontSize);

            const idxSize = (isScreen ? 24 : 12) * sc;
            let badgeX = x + 50*sc;
            let badgeY = y + 50*sc;
            let badgeTextX = x + 55*sc;
            let badgeTextY = y + 45*sc;
            let badgeRadius = 25*sc;
            if (currentTemplate === 6) { ctx.fillStyle = 'rgba(255,255,255,0.3)'; ctx.beginPath(); ctx.arc(badgeX, badgeY, badgeRadius, 0, 2 * Math.PI); ctx.fill(); ctx.fillStyle = '#ffffff'; }
            else if (currentTemplate === 4) {
                badgeX = x + w - 55*sc;
                badgeY = y + 50*sc;
                badgeTextX = badgeX;
                badgeTextY = y + 46*sc;
                badgeRadius = 27*sc;
                ctx.fillStyle = '#f5f5f7';
                ctx.beginPath();
                ctx.arc(badgeX, badgeY, badgeRadius, 0, 2 * Math.PI);
                ctx.fill();
                ctx.fillStyle = '#8e8e93';
            }
            else if(currentTemplate === 5) {
                 ctx.fillStyle = '#a1887f'; ctx.fillRect(x + 30*sc, y + 20*sc, 50*sc, 50*sc);
                 ctx.fillStyle = '#ffffff';
            }
            else if(currentTemplate === 2) ctx.fillStyle = UNI.green; else ctx.fillStyle = '#2c6060';
            if(currentTemplate !== 4 && currentTemplate !== 5) { ctx.beginPath(); ctx.arc(badgeX, badgeY, badgeRadius, 0, 2 * Math.PI); ctx.fill(); ctx.fillStyle = '#fff'; }
            ctx.font = `bold ${idxSize}px Cairo`; ctx.textAlign = 'center'; ctx.textBaseline = 'middle'; ctx.fillText(idx, badgeTextX, badgeTextY);
        }

        function getStats(a) { return { t: a.length, n: a.filter(x=>x.st==='جديدة').length, c: a.filter(x=>x.st==='مستمرة').length, i: a.filter(x=>x.st==='خارجية' || isExternalExecution(x.loc)).length }; }
        function chunk(a, s) { const r=[]; for(let i=0; i<a.length; i+=s) r.push(a.slice(i,i+s)); return r; }
