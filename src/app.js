// --- 1. Setup & Helpers ---
        let rows = [];
        let currentTemplate = 1;
        const ALLOWED_TEMPLATES = new Set([1, 2, 3, 4, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21]);
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
            room: ["LAB 1", "LAB 2", "LAB 3", "LAB 4", "LAB 5", "LAB 6", "CLASS 1", "CLASS 2", "CLASS 3", "CLASS 4", "CLASS 5", "CLASS 6", "CLASS 7", "CLASS 8", "CLASS 9", "CLASS 10", "مركز الجرائم السيبرانية", "مركز الذكاء الاصطناعي", "مركز السلامة المرورية", "كلية الادلة الجنائية", "النادي الرياضي", "عن بعد", "خارج المملكة", "خارج الرياض", "مركز التدريب - الامن العام"],
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

        function normalizeSupervisor(value, fallback = '') {
            const raw = String(value || '').trim();
            if (!raw) return fallback;
            const normalized = normalizeArabicValue(raw).replace(/\bال\s+/g, 'ال ');
            const matched = DDL.sup.find(name => normalizeArabicValue(name) === normalized);
            if (matched) return matched;
            const compact = normalized.replace(/\s+/g, '');
            const compactMatch = DDL.sup.find(name => normalizeArabicValue(name).replace(/\s+/g, '') === compact);
            return compactMatch || raw;
        }

        function getImportedValue(row, keys) {
            const wanted = Array.isArray(keys) ? keys : [keys];
            for (const key of wanted) {
                if (Object.prototype.hasOwnProperty.call(row, key) && row[key] !== undefined && row[key] !== null && String(row[key]).trim() !== '') {
                    return row[key];
                }
            }
            const normalizedWanted = wanted.map(key => normalizeArabicValue(key));
            const foundKey = Object.keys(row).find(key => {
                const normalizedKey = normalizeArabicValue(key);
                return normalizedWanted.includes(normalizedKey) ||
                    normalizedWanted.some(wantedKey => normalizedKey.includes(wantedKey) || wantedKey.includes(normalizedKey));
            });
            return foundKey ? row[foundKey] : '';
        }

        function normalizeExecutionPlace(value) {
            const raw = String(value || '').trim();
            const v = normalizeArabicValue(raw);
            if (!v) return DEFAULT_EXECUTION_PLACE;
            if (v.includes('جامعه نايف') || v.includes('نايف العربيه') || v.includes('مقر الجامعه') || v.includes('مقر الجامعة')) return 'مقر الجامعة';
            return raw;
        }

        function isRemoteValue(value) {
            const v = normalizeArabicValue(value);
            return v.includes('عن بعد') || v.includes('اونلاين') || v.includes('online') || v.includes('remote');
        }

        function isInternalExecutionPlace(value) {
            const v = normalizeArabicValue(value);
            if (!v) return true;
            return v.includes('جامعه نايف') || v.includes('جامعة نايف') || v.includes('نايف العربيه') || v.includes('مقر الجامعه') || v.includes('مقر الجامعة') || v === normalizeArabicValue(DEFAULT_EXECUTION_PLACE);
        }

        function applyExecutionPlaceRules(row) {
            if (!row) return row;
            row.loc = normalizeExecutionPlace(row.loc);
            if (isRemoteValue(row.st) || isRemoteValue(row.loc) || isRemoteValue(row.r) || isRemoteValue(row.f)) {
                row.r = 'عن بعد';
                row.f = 'عن بعد';
                if (isRemoteValue(row.loc)) row.loc = 'عن بعد';
                return row;
            }
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
            id = Number(id);
            if (!ALLOWED_TEMPLATES.has(id)) id = 1;
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
                sp: normalizeSupervisor(d.sp || d[KEYS.sp], lastUsed.sp || DDL.sup[0]),
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
            el.className = "course-row p-3 relative group";
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
                if (k === 'r') v = extractTrainingRoom(v);
                if (k === 'p') v = normalizePeriod(v);
                if (k === 'f') v = normalizeFloor(v);
                if (k === 'st') v = normalizeStatus(v);
                o[k] = v;
                if (k === 'loc' || k === 'st' || k === 'r' || k === 'f') {
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
                <button id="smart-paste-btn" type="button" onclick="openSmartPasteModal()" class="form-input flex items-center justify-center gap-1.5 text-sm hover:bg-slate-50">
                    <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="2" width="6" height="4" rx="1"/><path d="M17 4h1a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h1"/><path d="M9 12h6M9 16h4"/></svg>
                    لصق ذكي من LMS
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
                sp: normalizeSupervisor(clean[5], DDL.sup[0]),
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


        function isBlankScheduleRow(row) {
            if (!row) return true;
            return !String(row.n || '').trim() &&
                   !String(row.s || '').trim() &&
                   !String(row.e || '').trim();
        }

        function removeBlankDefaultRows() {
            rows = rows.filter(row => !isBlankScheduleRow(row));
            const box = document.getElementById('rows-box');
            if (box) {
                [...box.children].forEach(child => {
                    const input = child.querySelector('input[type="text"]');
                    const dateInputs = child.querySelectorAll('input[type="date"]');
                    const hasName = input && String(input.value || '').trim();
                    const hasStart = dateInputs[0] && String(dateInputs[0].value || '').trim();
                    const hasEnd = dateInputs[1] && String(dateInputs[1].value || '').trim();
                    if (!hasName && !hasStart && !hasEnd) child.remove();
                });
            }
            updateIdx();
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
            } else {
                removeBlankDefaultRows();
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
                    const json = XLSX.utils.sheet_to_json(sheet, { defval: '' });
                    document.getElementById('rows-box').innerHTML = '';
                    rows = [];
                    json.forEach(row => {
                        const importedSupervisor = getImportedValue(row, [KEYS.sp, 'منسق التدريب', 'المنسق', 'اسم المنسق', 'مسؤول التدريب']);
                        addRow({
                            n: getImportedValue(row, [KEYS.n, 'النشاط التدريبي', 'اسم الدورة', 'الدورة']) || '',
                            loc: normalizeExecutionPlace(getImportedValue(row, [KEYS.loc, 'مقر التنفيذ', 'مكان الدورة', 'المكان'])),
                            s: parseExcelDate(getImportedValue(row, [KEYS.s, 'بداية الدورة', 'البداية'])),
                            e: parseExcelDate(getImportedValue(row, [KEYS.e, 'نهاية الدورة', 'النهاية'])),
                            st: normalizeStatus(getImportedValue(row, [KEYS.st, 'حالة الدورة'])),
                            sp: normalizeSupervisor(importedSupervisor, DDL.sup[0]),
                            r: getImportedValue(row, [KEYS.r, 'قاعة التدريب', 'المعمل']) || DDL.room[0],
                            f: normalizeFloor(getImportedValue(row, [KEYS.f, 'الدور'])),
                            p: normalizePeriod(getImportedValue(row, [KEYS.p, 'النطاق', 'التوقيت']))
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
        function refreshArchiveView() {
            const container = document.getElementById('archive-list');
            const noData = document.getElementById('no-archive');
            const archive = JSON.parse(localStorage.getItem(ARCH_KEY) || '{}');
            const keys = Object.keys(archive);
            container.innerHTML = '';
            if (keys.length === 0) { noData.classList.remove('hidden'); return; }
            noData.classList.add('hidden');
            keys.sort((a, b) => b - a).forEach(key => {
                const entry = archive[key];
                const owner = entry.owner || 'غير محدد';
                const count = (entry.rows || entry).length;
                const div = document.createElement('div');
                div.className = "p-4 rounded-xl border";
                div.innerHTML = `
                    <div class="flex justify-between items-center mb-3">
                        <h4 class="font-black text-base" style="color:var(--brand-deep)">الأسبوع ${key}</h4>
                        <span class="badge badge-brand">${count} دورة</span>
                    </div>
                    <div class="text-xs mb-4" style="color:var(--muted)">
                        بواسطة: <span class="font-bold" style="color:var(--ink-sub)">${owner}</span>
                    </div>
                    <div class="flex gap-2">
                        <button onclick="loadFromArchive('${key}')" class="btn-main flex-1 text-xs gap-1.5">
                            <svg xmlns="http://www.w3.org/2000/svg" width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                            تعديل
                        </button>
                        <button onclick="deleteFromArchive('${key}')" class="btn-outline text-xs" style="color:var(--danger);border-color:rgba(143,47,45,0.3)" onmouseover="this.style.background='var(--danger-soft)'" onmouseout="this.style.background=''">
                            <svg xmlns="http://www.w3.org/2000/svg" width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"/><path d="m19 6-.867 12.142A2 2 0 0 1 16.138 20H7.862a2 2 0 0 1-1.995-1.858L5 6"/><path d="M10 11v6M14 11v6"/><path d="m8.5 6 .5-3h6l.5 3"/></svg>
                            حذف
                        </button>
                    </div>`;
                container.appendChild(div);
            });
        }

        // --- 5. Save Logic ---
        
        // اللوحة الأساسية الثابتة 2160×3840 (نسبة 9:16).
        // القوالب الاحترافية مصمّمة لتملأ هذه اللوحة بالكامل ثم تُصدّر بمضاعف ثابت:
        //   • القوالب 12–21 → ×2 = 4320×7680 (8K)، بلا قص ولا تجاوز ولا تخطيط متجاوب.
        //   • قالب iOS HD (4) → ×3 = 6480×11520.
        //   • باقي القوالب → ×1 = 2160×3840.
        const SCREEN_BASE_W = 2160, SCREEN_BASE_H = 3840;
        function getScreenExportScale() {
            if (currentTemplate === 4) return 3;
            if ([12,13,14,15,16,17,18,19,20,21].includes(currentTemplate)) return 2;
            return 1;
        }
        async function generateScreenZipBlob(valid, week) {
            const stats = getStats(valid);
            const zip = new JSZip();
            const chunksS = chunk(valid, 4);
            const scale = getScreenExportScale();
            for(let i=0; i<chunksS.length; i++) {
                const canvas = document.createElement('canvas');
                canvas.width  = SCREEN_BASE_W * scale;
                canvas.height = SCREEN_BASE_H * scale;
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
            if (currentTemplate === 4)  { drawIosHDScreenCard(ctx, list, stats, pn, tp, si, logo); return; }
            if (currentTemplate === 12) { drawReferenceOneScreenCard(ctx, list, stats, pn, tp, si, logo); return; }
            if (currentTemplate === 13) { drawReferenceTwoScreenCard(ctx, list, stats, pn, tp, si, logo); return; }
            if (currentTemplate === 14) { drawReferenceThreeScreenCard(ctx, list, stats, pn, tp, si, logo); return; }
            if (currentTemplate === 15) { drawReferenceFourScreenCard(ctx, list, stats, pn, tp, si, logo); return; }
            if (currentTemplate === 16) { drawReferenceFiveScreenCard(ctx, list, stats, pn, tp, si, logo); return; }
            if (currentTemplate === 17) { drawRef17ScreenCard(ctx, list, stats, pn, tp, si, logo); return; }
            if (currentTemplate === 18) { drawRef18ScreenCard(ctx, list, stats, pn, tp, si, logo); return; }
            if (currentTemplate === 19) { drawRef19ScreenCard(ctx, list, stats, pn, tp, si, logo); return; }
            if (currentTemplate === 20) { drawRef20ScreenCard(ctx, list, stats, pn, tp, si, logo); return; }
            if (currentTemplate === 21) { drawRef21ScreenCard(ctx, list, stats, pn, tp, si, logo); return; }
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

        function drawN8NScreenCard(ctx, list, stats, pn, tp, si, logo) {
            const W = 2160, H = 3840;
            ctx.fillStyle = '#f7faf8';
            ctx.fillRect(0, 0, W, H);

            ctx.fillStyle = UNI.green;
            ctx.fillRect(0, 0, W, 560);
            ctx.fillStyle = UNI.gold;
            ctx.fillRect(0, 540, W, 20);
            ctx.fillRect(W - 150, 0, 70, H);

            if (logo && logo.width > 0) {
                const logoW = 780;
                const logoH = logoW * (logo.height / logo.width);
                ctx.drawImage(logo, W - logoW - 210, 84, logoW, logoH);
            }

            ctx.fillStyle = '#ffffff';
            ctx.textAlign = 'right';
            ctx.font = 'normal 96px Cairo';
            ctx.fillText('جدول الدورات التدريبية', W - 220, 390);
            ctx.font = 'normal 44px Cairo';
            ctx.fillText(`صفحة ${pn} من ${tp}`, W - 220, 468);

            const introText = getScreenIntroText();
            ctx.fillStyle = '#2d4d4c';
            ctx.textAlign = 'center';
            ctx.font = 'normal 42px Cairo';
            let introY = 720;
            if (introText) {
                wrapTextSimple(ctx, introText, W - 300).slice(0, 3).forEach(line => { ctx.fillText(line, W / 2, introY); introY += 58; });
            }

            const statsY = introText ? introY + 30 : 700;
            drawN8NStats(ctx, stats, 210, statsY, W - 420);

            let y = statsY + 250;
            const cardGap = 34;
            const cardH = (H - y - 240 - (cardGap * 3)) / 4;
            for (let i = 0; i < list.length; i++) {
                if (list[i]) drawN8NCourseCard(ctx, list[i], 190, y, W - 420, cardH, si + i + 1);
                y += cardH + cardGap;
            }

            ctx.fillStyle = '#e8eeed';
            ctx.fillRect(0, H - 130, W, 130);
            ctx.fillStyle = UNI.green;
            ctx.textAlign = 'center';
            ctx.font = 'normal 34px Cairo';
            ctx.fillText('وكالة التدريب بجامعة نايف العربية للعلوم الأمنية', W / 2, H - 48);
        }

        function drawN8NStats(ctx, stats, x, y, w) {
            const labels = [
                ['إجمالي', stats.t],
                ['جديدة', stats.n],
                ['مستمرة', stats.c],
                ['خارجية', stats.i]
            ];
            const gap = 24;
            const boxW = (w - gap * 3) / 4;
            labels.forEach((item, index) => {
                const bx = x + index * (boxW + gap);
                ctx.fillStyle = index === 0 ? UNI.green : '#ffffff';
                drawRoundedRect(ctx, bx, y, boxW, 170, 18);
                ctx.strokeStyle = index === 0 ? UNI.green : '#d8dedb';
                ctx.lineWidth = 3;
                ctx.strokeRect(bx, y, boxW, 170);
                ctx.fillStyle = index === 0 ? '#ffffff' : UNI.green;
                ctx.textAlign = 'center';
                ctx.textBaseline = 'middle';
                ctx.font = 'normal 72px Cairo';
                ctx.fillText(item[1], bx + boxW / 2, y + 66);
                ctx.font = 'normal 34px Cairo';
                ctx.fillText(item[0], bx + boxW / 2, y + 125);
            });
        }

        function drawN8NCourseCard(ctx, c, x, y, w, h, idx) {
            ctx.fillStyle = '#ffffff';
            ctx.shadowColor = 'rgba(1, 101, 100, 0.10)';
            ctx.shadowBlur = 26;
            ctx.shadowOffsetY = 8;
            drawRoundedRect(ctx, x, y, w, h, 22);
            ctx.shadowColor = 'transparent';
            ctx.fillStyle = UNI.gold;
            ctx.fillRect(x + w - 24, y, 24, h);

            ctx.fillStyle = UNI.green;
            ctx.textAlign = 'right';
            ctx.font = 'normal 68px Cairo';
            const title = wrapTextSmart(ctx, c.n + (isExternalExecution(c.loc) ? '  🌐' : ''), w - 230, 2, 68);
            let titleY = y + 86;
            title.lines.forEach(line => { ctx.fillText(line, x + w - 80, titleY); titleY += title.lineHeight; });

            ctx.fillStyle = '#5d6f6d';
            ctx.font = 'normal 38px Cairo';
            ctx.fillText(`${c.s} - ${c.e}`, x + w - 80, y + h - 76);

            const info = [
                ['الحالة', c.st],
                ['الفترة', c.p],
                ['المكان', c.loc],
                ['القاعة', c.r],
                ['الدور', shouldHideFloorValue(c.f) ? '-' : c.f]
            ];
            const boxGap = 18;
            const boxW = (w - 210 - boxGap * 4) / 5;
            let bx = x + 62;
            const by = y + h - 190;
            info.forEach(item => {
                ctx.fillStyle = '#f3f7f6';
                drawRoundedRect(ctx, bx, by, boxW, 108, 14);
                ctx.fillStyle = '#6f807d';
                ctx.textAlign = 'center';
                ctx.font = 'normal 23px Cairo';
                ctx.fillText(item[0], bx + boxW / 2, by + 34);
                ctx.fillStyle = UNI.green;
                ctx.font = 'normal 30px Cairo';
                let value = String(item[1] || '-');
                while (ctx.measureText(value).width > boxW - 24 && value.length > 0) value = value.slice(0, -1);
                if (value !== String(item[1] || '-')) value += '...';
                ctx.fillText(value, bx + boxW / 2, by + 78);
                bx += boxW + boxGap;
            });

            ctx.fillStyle = 'rgba(128,47,45,0.08)';
            drawRoundedRect(ctx, x + 48, y + 42, 600, 96, 18);
            ctx.fillStyle = UNI.supRed;
            ctx.textAlign = 'left';
            ctx.font = 'normal 34px Cairo';
            ctx.fillText(`منسق التدريب: ${c.sp}`, x + 70, y + 72);
            ctx.fillStyle = UNI.green;
            ctx.font = 'normal 34px Cairo';
            ctx.fillText(String(idx).padStart(2, '0'), x + 72, y + 126);
        }

        function drawN9NScreenCard(ctx, list, stats, pn, tp, si, logo) {
            const W = 2160, H = 3840;
            const grd = ctx.createLinearGradient(0, 0, W, H);
            grd.addColorStop(0, '#012f2f');
            grd.addColorStop(1, '#016564');
            ctx.fillStyle = grd;
            ctx.fillRect(0, 0, W, H);

            ctx.strokeStyle = 'rgba(208, 178, 132, 0.38)';
            ctx.lineWidth = 4;
            ctx.strokeRect(70, 70, W - 140, H - 140);
            ctx.fillStyle = 'rgba(208, 178, 132, 0.12)';
            ctx.fillRect(112, 112, W - 224, 8);
            ctx.fillRect(112, H - 120, W - 224, 8);

            if (logo && logo.width > 0) {
                const logoW = 720;
                const logoH = logoW * (logo.height / logo.width);
                ctx.drawImage(logo, (W - logoW) / 2, 125, logoW, logoH);
            }

            ctx.fillStyle = '#ffffff';
            ctx.textAlign = 'center';
            ctx.font = 'normal 92px Cairo';
            ctx.fillText('جدول الدورات التدريبية', W / 2, 455);
            ctx.fillStyle = UNI.gold;
            ctx.font = 'normal 36px Cairo';
            ctx.fillText(`صفحة ${pn} من ${tp}`, W / 2, 520);

            drawN9NStats(ctx, stats, 185, 650, W - 370);

            const introText = getScreenIntroText();
            ctx.fillStyle = 'rgba(255,255,255,0.92)';
            ctx.font = 'normal 40px Cairo';
            let introY = 980;
            if (introText) {
                wrapTextSimple(ctx, introText, W - 330).slice(0, 3).forEach(line => { ctx.fillText(line, W / 2, introY); introY += 58; });
            }

            let y = introText ? introY + 70 : 1000;
            const cardGap = 32;
            const cardH = (H - y - 230 - (cardGap * 3)) / 4;
            for (let i = 0; i < list.length; i++) {
                if (list[i]) drawN9NCourseCard(ctx, list[i], 165, y, W - 330, cardH, si + i + 1);
                y += cardH + cardGap;
            }

            ctx.fillStyle = 'rgba(0,0,0,0.18)';
            ctx.fillRect(70, H - 190, W - 140, 90);
            ctx.fillStyle = '#ffffff';
            ctx.font = 'normal 32px Cairo';
            ctx.fillText('وكالة التدريب بجامعة نايف العربية للعلوم الأمنية', W / 2, H - 132);
        }

        function drawN9NStats(ctx, stats, x, y, w) {
            const items = [['إجمالي', stats.t], ['جديدة', stats.n], ['مستمرة', stats.c], ['خارجية', stats.i]];
            const gap = 22;
            const boxW = (w - gap * 3) / 4;
            items.forEach((item, index) => {
                const bx = x + index * (boxW + gap);
                ctx.fillStyle = index === 0 ? UNI.gold : 'rgba(255,255,255,0.08)';
                drawRoundedRect(ctx, bx, y, boxW, 190, 8);
                ctx.strokeStyle = 'rgba(208, 178, 132, 0.55)';
                ctx.lineWidth = 2;
                ctx.strokeRect(bx, y, boxW, 190);
                ctx.fillStyle = index === 0 ? '#013c3b' : '#ffffff';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'middle';
                ctx.font = 'normal 78px Cairo';
                ctx.fillText(item[1], bx + boxW / 2, y + 76);
                ctx.font = 'normal 35px Cairo';
                ctx.fillText(item[0], bx + boxW / 2, y + 137);
            });
        }

        function drawN9NCourseCard(ctx, c, x, y, w, h, idx) {
            ctx.fillStyle = 'rgba(255, 255, 255, 0.96)';
            drawRoundedRect(ctx, x, y, w, h, 8);
            ctx.fillStyle = UNI.gold;
            ctx.fillRect(x, y, w, 10);

            ctx.fillStyle = UNI.green;
            ctx.textAlign = 'right';
            ctx.font = 'normal 66px Cairo';
            const result = wrapTextSmart(ctx, c.n + (isExternalExecution(c.loc) ? '  🌐' : ''), w - 230, 2, 66);
            let ty = y + 88;
            result.lines.forEach(line => { ctx.fillText(line, x + w - 70, ty); ty += result.lineHeight; });

            ctx.fillStyle = '#526461';
            ctx.font = 'normal 38px Cairo';
            ctx.fillText(`${c.s} - ${c.e}`, x + w - 70, y + h - 68);

            const details = [
                `${c.st}`,
                `${c.p}`,
                `${c.loc}`,
                `${c.r}`,
                `${shouldHideFloorValue(c.f) ? '-' : c.f}`
            ];
            ctx.textAlign = 'center';
            const gap = 16;
            const pillW = (w - 240 - gap * 4) / 5;
            let px = x + 70;
            const py = y + h - 178;
            details.forEach(value => {
                ctx.fillStyle = '#edf3f1';
                drawRoundedRect(ctx, px, py, pillW, 92, 46);
                ctx.fillStyle = UNI.green;
                ctx.font = 'normal 31px Cairo';
                let safe = String(value || '-');
                while (ctx.measureText(safe).width > pillW - 32 && safe.length > 0) safe = safe.slice(0, -1);
                if (safe !== String(value || '-')) safe += '...';
                ctx.fillText(safe, px + pillW / 2, py + 58);
                px += pillW + gap;
            });

            ctx.fillStyle = UNI.gold;
            ctx.textAlign = 'left';
            ctx.font = 'normal 38px Cairo';
            ctx.fillText(String(idx).padStart(2, '0'), x + 70, y + 70);
            ctx.fillStyle = 'rgba(128,47,45,0.08)';
            drawRoundedRect(ctx, x + 60, y + 84, 560, 74, 16);
            ctx.fillStyle = UNI.supRed;
            ctx.font = 'normal 33px Cairo';
            ctx.fillText(`منسق التدريب: ${c.sp}`, x + 70, y + 118);
        }

        function drawOpsScreenCard(ctx, list, stats, pn, tp, si, logo) {
            const W = 2160, H = 3840;
            ctx.fillStyle = '#f3f6f4';
            ctx.fillRect(0, 0, W, H);

            ctx.fillStyle = '#123f3d';
            ctx.fillRect(W - 470, 0, 470, H);
            ctx.fillStyle = UNI.gold;
            ctx.fillRect(W - 500, 0, 30, H);
            ctx.fillStyle = 'rgba(255,255,255,0.08)';
            for (let y = 180; y < H; y += 260) ctx.fillRect(W - 430, y, 310, 2);

            if (logo && logo.width > 0) {
                const logoW = 330;
                const logoH = logoW * (logo.height / logo.width);
                ctx.drawImage(logo, W - 400, 115, logoW, logoH);
            }

            ctx.save();
            ctx.translate(W - 250, 1060);
            ctx.rotate(-Math.PI / 2);
            ctx.fillStyle = '#ffffff';
            ctx.textAlign = 'center';
            ctx.font = 'normal 86px Cairo';
            ctx.fillText('جدول الدورات التدريبية', 0, 0);
            ctx.restore();

            ctx.fillStyle = UNI.gold;
            ctx.textAlign = 'center';
            ctx.font = 'normal 38px Cairo';
            ctx.fillText(`صفحة ${pn} من ${tp}`, W - 250, 1520);
            drawOpsStats(ctx, stats, W - 405, 1780, 340);

            const introText = getScreenIntroText();
            ctx.fillStyle = '#365755';
            ctx.textAlign = 'right';
            ctx.font = 'normal 42px Cairo';
            let introY = 190;
            if (introText) {
                wrapTextSimple(ctx, introText, 1500).slice(0, 3).forEach(line => {
                    ctx.fillText(line, W - 610, introY);
                    introY += 58;
                });
            }

            const gridX = 130;
            const gridY = introText ? 430 : 280;
            const gap = 42;
            const cardW = (W - 690 - gap) / 2;
            const cardH = (H - gridY - 240 - gap) / 2;
            for (let i = 0; i < list.length; i++) {
                const col = i % 2;
                const row = Math.floor(i / 2);
                if (list[i]) drawOpsCourseCard(ctx, list[i], gridX + col * (cardW + gap), gridY + row * (cardH + gap), cardW, cardH, si + i + 1);
            }

            ctx.fillStyle = '#dfe7e4';
            ctx.fillRect(0, H - 128, W - 500, 128);
            ctx.fillStyle = '#123f3d';
            ctx.textAlign = 'center';
            ctx.font = 'normal 34px Cairo';
            ctx.fillText('وكالة التدريب بجامعة نايف العربية للعلوم الأمنية', (W - 500) / 2, H - 48);
        }

        function drawOpsStats(ctx, stats, x, y, w) {
            const items = [['إجمالي', stats.t], ['جديدة', stats.n], ['مستمرة', stats.c], ['خارجية', stats.i]];
            items.forEach((item, index) => {
                const by = y + index * 210;
                ctx.fillStyle = index === 0 ? UNI.gold : 'rgba(255,255,255,0.10)';
                drawRoundedRect(ctx, x, by, w, 165, 18);
                ctx.fillStyle = index === 0 ? '#123f3d' : '#ffffff';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'middle';
                ctx.font = 'normal 70px Cairo';
                ctx.fillText(item[1], x + w / 2, by + 62);
                ctx.font = 'normal 32px Cairo';
                ctx.fillText(item[0], x + w / 2, by + 118);
            });
        }

        function drawOpsCourseCard(ctx, c, x, y, w, h, idx) {
            ctx.fillStyle = '#ffffff';
            ctx.shadowColor = 'rgba(18,63,61,0.14)';
            ctx.shadowBlur = 30;
            ctx.shadowOffsetY = 10;
            drawRoundedRect(ctx, x, y, w, h, 26);
            ctx.shadowColor = 'transparent';
            ctx.fillStyle = '#123f3d';
            ctx.fillRect(x, y, w, 86);
            ctx.fillStyle = UNI.gold;
            ctx.fillRect(x + w - 18, y, 18, h);

            ctx.fillStyle = '#ffffff';
            ctx.textAlign = 'left';
            ctx.font = 'normal 34px Cairo';
            ctx.fillText(String(idx).padStart(2, '0'), x + 45, y + 56);

            ctx.fillStyle = '#123f3d';
            ctx.textAlign = 'right';
            const title = wrapTextSmart(ctx, c.n + (isExternalExecution(c.loc) ? '  🌐' : ''), w - 95, 3, 58);
            let ty = y + 190;
            ctx.font = `normal ${title.fontSize}px Cairo`;
            title.lines.forEach(line => {
                ctx.fillText(line, x + w - 55, ty);
                ty += title.lineHeight;
            });

            ctx.fillStyle = '#637572';
            ctx.font = 'normal 35px Cairo';
            ctx.fillText(`${c.s} - ${c.e}`, x + w - 55, y + h - 236);

            const info = [['الحالة', c.st], ['الفترة', c.p], ['المكان', c.loc], ['القاعة', c.r], ['الدور', shouldHideFloorValue(c.f) ? '-' : c.f]];
            const boxGap = 14;
            const boxW = (w - 96 - boxGap * 4) / 5;
            let bx = x + 40;
            const by = y + h - 170;
            info.forEach(item => {
                ctx.fillStyle = '#eef4f2';
                drawRoundedRect(ctx, bx, by, boxW, 96, 14);
                ctx.fillStyle = '#6b7d79';
                ctx.textAlign = 'center';
                ctx.font = 'normal 19px Cairo';
                ctx.fillText(item[0], bx + boxW / 2, by + 31);
                ctx.fillStyle = '#123f3d';
                ctx.font = 'normal 26px Cairo';
                let value = String(item[1] || '-');
                while (ctx.measureText(value).width > boxW - 18 && value.length > 0) value = value.slice(0, -1);
                if (value !== String(item[1] || '-')) value += '...';
                ctx.fillText(value, bx + boxW / 2, by + 70);
                bx += boxW + boxGap;
            });

            ctx.fillStyle = 'rgba(128,47,45,0.08)';
            drawRoundedRect(ctx, x + 42, y + 105, 520, 70, 14);
            ctx.fillStyle = UNI.supRed;
            ctx.textAlign = 'left';
            ctx.font = 'normal 30px Cairo';
            ctx.fillText(`منسق التدريب: ${c.sp}`, x + 62, y + 151);
        }

        function drawTimelineScreenCard(ctx, list, stats, pn, tp, si, logo) {
            const W = 2160, H = 3840;
            ctx.fillStyle = '#fffaf0';
            ctx.fillRect(0, 0, W, H);
            ctx.fillStyle = '#f2e6cb';
            ctx.fillRect(0, 0, W, 18);
            ctx.fillRect(0, H - 18, W, 18);

            if (logo && logo.width > 0) {
                const logoW = 720;
                const logoH = logoW * (logo.height / logo.width);
                ctx.drawImage(logo, W - logoW - 110, 95, logoW, logoH);
            }

            ctx.fillStyle = '#7a5a24';
            ctx.textAlign = 'right';
            ctx.font = 'normal 92px Cairo';
            ctx.fillText('جدول الدورات التدريبية', W - 120, 430);
            ctx.fillStyle = '#016564';
            ctx.font = 'normal 38px Cairo';
            ctx.fillText(`صفحة ${pn} من ${tp}`, W - 120, 500);
            drawTimelineStats(ctx, stats, 120, 120, 820);

            const introText = getScreenIntroText();
            ctx.fillStyle = '#4b5f5c';
            ctx.textAlign = 'center';
            ctx.font = 'normal 40px Cairo';
            let introY = 675;
            if (introText) {
                wrapTextSimple(ctx, introText, W - 260).slice(0, 3).forEach(line => {
                    ctx.fillText(line, W / 2, introY);
                    introY += 58;
                });
            }

            const timelineX = W / 2;
            const startY = introText ? introY + 80 : 720;
            const cardH = 500;
            const gap = 60;
            ctx.strokeStyle = '#d0b284';
            ctx.lineWidth = 8;
            ctx.beginPath();
            ctx.moveTo(timelineX, startY - 40);
            ctx.lineTo(timelineX, H - 245);
            ctx.stroke();

            for (let i = 0; i < list.length; i++) {
                const y = startY + i * (cardH + gap);
                if (list[i]) drawTimelineCourseCard(ctx, list[i], timelineX, y, cardH, si + i + 1, i % 2 === 0);
            }

            ctx.fillStyle = '#016564';
            ctx.fillRect(0, H - 145, W, 145);
            ctx.fillStyle = UNI.gold;
            ctx.fillRect(0, H - 145, W, 10);
            ctx.fillStyle = '#ffffff';
            ctx.textAlign = 'center';
            ctx.font = 'normal 34px Cairo';
            ctx.fillText('وكالة التدريب بجامعة نايف العربية للعلوم الأمنية', W / 2, H - 56);
        }

        function drawTimelineStats(ctx, stats, x, y, w) {
            const items = [['إجمالي', stats.t], ['جديدة', stats.n], ['مستمرة', stats.c], ['خارجية', stats.i]];
            const gap = 18;
            const boxW = (w - gap * 3) / 4;
            items.forEach((item, index) => {
                const bx = x + index * (boxW + gap);
                ctx.fillStyle = index === 0 ? '#016564' : '#ffffff';
                drawRoundedRect(ctx, bx, y, boxW, 145, 18);
                ctx.strokeStyle = '#d8c59e';
                ctx.lineWidth = 2;
                ctx.strokeRect(bx, y, boxW, 145);
                ctx.fillStyle = index === 0 ? '#ffffff' : '#7a5a24';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'middle';
                ctx.font = 'normal 56px Cairo';
                ctx.fillText(item[1], bx + boxW / 2, y + 54);
                ctx.font = 'normal 25px Cairo';
                ctx.fillText(item[0], bx + boxW / 2, y + 104);
            });
        }

        function drawTimelineCourseCard(ctx, c, timelineX, y, h, idx, rightSide) {
            const cardW = 830;
            const connector = 95;
            const x = rightSide ? timelineX + connector : timelineX - connector - cardW;

            ctx.fillStyle = '#016564';
            ctx.beginPath();
            ctx.arc(timelineX, y + h / 2, 28, 0, Math.PI * 2);
            ctx.fill();
            ctx.fillStyle = UNI.gold;
            ctx.beginPath();
            ctx.arc(timelineX, y + h / 2, 13, 0, Math.PI * 2);
            ctx.fill();
            ctx.strokeStyle = UNI.gold;
            ctx.lineWidth = 5;
            ctx.beginPath();
            ctx.moveTo(timelineX, y + h / 2);
            ctx.lineTo(rightSide ? x : x + cardW, y + h / 2);
            ctx.stroke();

            ctx.fillStyle = '#ffffff';
            ctx.shadowColor = 'rgba(122,90,36,0.14)';
            ctx.shadowBlur = 24;
            ctx.shadowOffsetY = 8;
            drawRoundedRect(ctx, x, y, cardW, h, 24);
            ctx.shadowColor = 'transparent';

            ctx.fillStyle = rightSide ? '#016564' : '#7a5a24';
            ctx.fillRect(x, y, cardW, 78);
            ctx.fillStyle = '#ffffff';
            ctx.textAlign = rightSide ? 'left' : 'right';
            ctx.font = 'normal 32px Cairo';
            ctx.fillText(String(idx).padStart(2, '0'), rightSide ? x + 38 : x + cardW - 38, y + 51);

            ctx.fillStyle = '#1f3f3d';
            ctx.textAlign = 'right';
            const result = wrapTextSmart(ctx, c.n + (isExternalExecution(c.loc) ? '  🌐' : ''), cardW - 80, 2, 50);
            ctx.font = `normal ${result.fontSize}px Cairo`;
            let ty = y + 150;
            result.lines.forEach(line => {
                ctx.fillText(line, x + cardW - 40, ty);
                ty += result.lineHeight;
            });

            ctx.fillStyle = '#6b6254';
            ctx.font = 'normal 31px Cairo';
            ctx.fillText(`${c.s} - ${c.e}`, x + cardW - 40, y + 270);

            const chips = [c.st, c.p, c.loc, c.r, shouldHideFloorValue(c.f) ? '-' : c.f];
            const chipGap = 12;
            const chipW = (cardW - 80 - chipGap * 2) / 3;
            chips.slice(0, 3).forEach((value, i) => drawTimelineChip(ctx, String(value || '-'), x + 40 + i * (chipW + chipGap), y + 318, chipW));
            chips.slice(3).forEach((value, i) => drawTimelineChip(ctx, String(value || '-'), x + 40 + i * (chipW + chipGap), y + 392, chipW));

            ctx.fillStyle = 'rgba(128,47,45,0.08)';
            drawRoundedRect(ctx, x + 40, y + h - 68, cardW - 80, 48, 12);
            ctx.fillStyle = UNI.supRed;
            ctx.textAlign = 'center';
            ctx.font = 'normal 26px Cairo';
            ctx.fillText(`منسق التدريب: ${c.sp}`, x + cardW / 2, y + h - 36);
        }

        function drawTimelineChip(ctx, value, x, y, w) {
            ctx.fillStyle = '#f5efe0';
            drawRoundedRect(ctx, x, y, w, 56, 28);
            ctx.fillStyle = '#016564';
            ctx.textAlign = 'center';
            ctx.font = 'normal 24px Cairo';
            let safe = value;
            while (ctx.measureText(safe).width > w - 26 && safe.length > 0) safe = safe.slice(0, -1);
            if (safe !== value) safe += '...';
            ctx.fillText(safe, x + w / 2, y + 37);
        }

        // =========================================================
        // 5 PROFESSIONAL TEMPLATES — 8K Portrait (4320 × 7680)
        // قوالب احترافية جديدة — 8K عمودي
        // =========================================================

        const T_W = 4320, T_H = 7680, T_MX = 155;
        const TC = {
            green:'#2A6364', greenD:'#1B4546', greenL:'#3C7778',
            gold:'#C7B08C', goldD:'#A89068', goldL:'rgba(199,176,140,0.38)',
            white:'#F9F9F9', bg:'#F1F5F3', bgCard:'#FFFFFF',
            ink:'#192C2C', sub:'#466768', muted:'#7A9898',
            soft:'rgba(42,99,100,0.07)', line:'rgba(42,99,100,0.11)',
        };
        function tSh(ctx,b,o,c){ctx.shadowColor=c||'rgba(26,68,69,0.13)';ctx.shadowBlur=b||44;ctx.shadowOffsetY=o||15;ctx.shadowOffsetX=0;}
        function tCl(ctx){ctx.shadowColor='transparent';ctx.shadowBlur=0;ctx.shadowOffsetY=0;}

        // ── Page background with dot texture ─────────────────────
        function tBg(ctx) {
            ctx.fillStyle = TC.bg; ctx.fillRect(0,0,T_W,T_H);
            ctx.save(); ctx.globalAlpha=0.026; ctx.fillStyle=TC.green;
            for(let xi=0;xi<=T_W;xi+=76) for(let yi=0;yi<=T_H;yi+=76){ctx.beginPath();ctx.arc(xi,yi,2.8,0,Math.PI*2);ctx.fill();}
            ctx.restore();
        }

        // ── Header: logo + gold lines + title ────────────────────
        function tHdr(ctx, logo, pn, tp) {
            let y = 135;
            if(logo&&logo.width>0){const lW=1500,lH=lW*(logo.height/logo.width);ctx.drawImage(logo,(T_W-lW)/2,y,lW,lH);y+=lH+58;}
            else{y+=390;}
            // ornament line above title
            ctx.fillStyle=TC.gold; ctx.fillRect((T_W-760)/2,y,760,7); y+=60;
            ctx.fillStyle=TC.green; ctx.textAlign='center'; ctx.font='bold 220px Cairo';
            ctx.fillText('جدول الدورات التدريبية',T_W/2,y+212); y+=278;
            ctx.fillStyle=TC.gold; ctx.fillRect((T_W-760)/2,y,760,7); y+=52;
            return y;
        }

        // ── Stats block RTL: خارجية، مستمرة، جديدة، إجمالي ──────
        function tSts(ctx, stats, y) {
            const items=[{l:'خارجية',v:stats.i,ic:0},{l:'مستمرة',v:stats.c,ic:1},{l:'جديدة',v:stats.n,ic:2},{l:'إجمالي',v:stats.t,ic:3}];
            const bH=335,gap=66,mX=T_MX+55,bW=(T_W-mX*2-gap*3)/4;
            items.forEach((it,i)=>{
                const bx=mX+i*(bW+gap);
                tSh(ctx,32,11,'rgba(26,68,69,0.09)'); ctx.fillStyle=TC.bgCard; drawRoundedRect(ctx,bx,y,bW,bH,28); tCl(ctx);
                ctx.strokeStyle=i===3?TC.gold:TC.line; ctx.lineWidth=i===3?5:3; ctx.strokeRect(bx,y,bW,bH);
                if(i===3){ctx.fillStyle=TC.green;ctx.fillRect(bx,y,bW,14);}
                tStIc(ctx,it.ic,bx+bW/2,y+85,44);
                ctx.fillStyle=i===3?TC.green:TC.ink; ctx.textAlign='center'; ctx.textBaseline='middle';
                ctx.font='bold 140px Cairo'; ctx.fillText(String(it.v),bx+bW/2,y+bH*0.52);
                ctx.fillStyle=TC.sub; ctx.font='normal 62px Cairo'; ctx.fillText(it.l,bx+bW/2,y+bH*0.82);
                ctx.textBaseline='alphabetic';
            });
            return y+bH;
        }
        function tStIc(ctx,t,cx,cy,s){
            ctx.save(); ctx.strokeStyle=TC.gold; ctx.lineWidth=5; ctx.lineCap='round';
            if(t===0){ctx.beginPath();ctx.arc(cx,cy,s,0,Math.PI*2);ctx.stroke();ctx.beginPath();ctx.moveTo(cx-s,cy);ctx.lineTo(cx+s,cy);ctx.stroke();ctx.beginPath();ctx.ellipse(cx,cy,s*.5,s,0,0,Math.PI*2);ctx.stroke();}
            else if(t===1){ctx.beginPath();ctx.arc(cx,cy,s*.8,-Math.PI*.5,Math.PI,false);ctx.stroke();ctx.beginPath();ctx.arc(cx,cy,s*.8,Math.PI,Math.PI*2.5,false);ctx.stroke();}
            else if(t===2){for(let k=0;k<5;k++){const a=k*Math.PI*2/5-Math.PI/2;ctx.beginPath();ctx.moveTo(cx+Math.cos(a)*s,cy+Math.sin(a)*s);ctx.lineTo(cx+Math.cos(a+Math.PI/5)*s*.42,cy+Math.sin(a+Math.PI/5)*s*.42);ctx.stroke();}}
            else{ctx.beginPath();ctx.arc(cx,cy-s*.12,s*.40,0,Math.PI*2);ctx.stroke();ctx.beginPath();ctx.arc(cx,cy,s*.78,Math.PI,Math.PI*2);ctx.stroke();}
            ctx.restore();
        }

        // ── Intro text box ────────────────────────────────────────
        function tIntr(ctx, y) {
            const txt=getScreenIntroText(); if(!txt) return y;
            ctx.font='normal 66px Cairo';
            const lines=wrapTextSimple(ctx,txt,T_W-640).slice(0,2);
            const bH=lines.length*106+80;
            ctx.fillStyle=TC.soft; drawRoundedRect(ctx,T_MX+55,y,T_W-(T_MX+55)*2,bH,22);
            ctx.strokeStyle=TC.goldL; ctx.lineWidth=3; ctx.strokeRect(T_MX+55,y,T_W-(T_MX+55)*2,bH);
            ctx.fillStyle=TC.sub; ctx.textAlign='center';
            let ty=y+78; lines.forEach(l=>{ctx.fillText(l,T_W/2,ty);ty+=106;});
            return y+bH+55;
        }

        // ── Footer ───────────────────────────────────────────────
        function tFtr(ctx) {
            const fH=178,fY=T_H-fH-28;
            const g=ctx.createLinearGradient(0,fY,0,T_H); g.addColorStop(0,TC.green); g.addColorStop(1,TC.greenD);
            ctx.fillStyle=g; drawRoundedRect(ctx,T_MX,fY,T_W-T_MX*2,fH+10,28);
            ctx.fillStyle=TC.gold; ctx.fillRect(T_MX,fY,T_W-T_MX*2,10);
            ctx.fillStyle='#FFFFFF'; ctx.textAlign='center'; ctx.font='bold 70px Cairo';
            ctx.fillText('وكالة التدريب بجامعة نايف العربية للعلوم الأمنية',T_W/2,fY+112);
            const dm=(x,y)=>{ctx.fillStyle=TC.gold;ctx.save();ctx.translate(x,y);ctx.rotate(Math.PI/4);ctx.fillRect(-13,-13,26,26);ctx.restore();};
            dm(T_W/2-680,fY+90); dm(T_W/2+680,fY+90);
        }

        // ── 5 info chips RTL (right=الطابق → left=مكان التنفيذ) ─
        // ─── Chip icons (gold thin lines) ─────────────────────────
        function tChipIcon(ctx, type, cx, cy, s) {
            ctx.save(); ctx.strokeStyle=TC.gold;
            ctx.lineWidth=Math.max(5,s*0.10); ctx.lineCap='round'; ctx.lineJoin='round';
            if(type===0){
                ctx.beginPath(); ctx.arc(cx,cy-s*0.12,s*0.38,0,Math.PI*2); ctx.stroke();
                ctx.beginPath(); ctx.moveTo(cx-s*0.38,cy-s*0.12);
                ctx.quadraticCurveTo(cx-s*0.52,cy+s*0.30,cx,cy+s*0.62);
                ctx.quadraticCurveTo(cx+s*0.52,cy+s*0.30,cx+s*0.38,cy-s*0.12); ctx.stroke();
            } else if(type===1){
                ctx.beginPath(); ctx.arc(cx,cy,s*0.50,0,Math.PI*2); ctx.stroke();
                ctx.beginPath(); ctx.moveTo(cx,cy); ctx.lineTo(cx,cy-s*0.28); ctx.stroke();
                ctx.beginPath(); ctx.moveTo(cx,cy); ctx.lineTo(cx+s*0.22,cy+s*0.14); ctx.stroke();
            } else if(type===2){
                ctx.beginPath(); ctx.arc(cx,cy,s*0.50,0,Math.PI*2); ctx.stroke();
                ctx.beginPath(); ctx.moveTo(cx-s*0.22,cy+s*0.02);
                ctx.lineTo(cx-s*0.04,cy+s*0.22); ctx.lineTo(cx+s*0.28,cy-s*0.20); ctx.stroke();
            } else if(type===3){
                ctx.beginPath(); ctx.rect(cx-s*0.30,cy-s*0.45,s*0.60,s*0.90); ctx.stroke();
                ctx.beginPath(); ctx.arc(cx+s*0.10,cy+s*0.05,s*0.07,0,Math.PI*2); ctx.stroke();
            } else {
                ctx.beginPath();
                ctx.moveTo(cx-s*0.42,cy+s*0.48); ctx.lineTo(cx-s*0.42,cy-s*0.48);
                ctx.lineTo(cx+s*0.42,cy-s*0.48); ctx.lineTo(cx+s*0.42,cy+s*0.48); ctx.stroke();
                [-s*0.16, s*0.10].forEach(oy=>{
                    ctx.beginPath(); ctx.moveTo(cx-s*0.42,cy+oy); ctx.lineTo(cx+s*0.42,cy+oy); ctx.stroke();
                });
            }
            ctx.restore();
        }

        // ─── Big chip row — sizes proportional to h (chip height) ─
        function tChip(ctx, c, x, y, w, h) {
            const it=[
                ['مكان التنفيذ',c.loc,0],['الفترة',c.p,1],
                ['الحالة',c.st,2],['القاعة',c.r,3],
                ['الطابق',shouldHideFloorValue(c.f)?'—':c.f,4],
            ];
            const gap=24, bW=(w-gap*4)/5;
            const iconS = h*0.20;
            const valFs = Math.min(190, h*0.35);
            const lblFs = Math.min(74, h*0.16);
            const iconCY = y+h*0.28;
            const valY   = y+h*0.64;
            const lblY   = y+h*0.86;
            it.forEach((item,i)=>{
                const bx=x+i*(bW+gap);
                ctx.fillStyle=TC.soft; drawRoundedRect(ctx,bx,y,bW,h,24);
                ctx.strokeStyle=TC.goldL; ctx.lineWidth=3; ctx.strokeRect(bx,y,bW,h);
                ctx.fillStyle=TC.gold; ctx.fillRect(bx,y,bW,7);
                tChipIcon(ctx, item[2], bx+bW/2, iconCY, iconS);
                ctx.fillStyle=TC.ink; ctx.textAlign='center'; ctx.font=`bold ${valFs}px Cairo`;
                let v=String(item[1]||'—');
                while(ctx.measureText(v).width>bW-44&&v.length>0) v=v.slice(0,-1);
                if(v!==String(item[1]||'—')) v+='...';
                ctx.fillText(v, bx+bW/2, valY);
                ctx.fillStyle=TC.muted; ctx.font=`normal ${lblFs}px Cairo`;
                ctx.fillText(item[0], bx+bW/2, lblY);
            });
        }

        // ─── Title 2-line max, proportional ───────────────────────
        function tTtl(ctx, text, maxW, fs, cx, y) {
            const r=wrapTextSmart(ctx,text,maxW,2,fs);
            ctx.font=`bold ${r.fontSize}px Cairo`;
            r.lines.forEach((l,i)=>ctx.fillText(l,cx,y+i*r.lineHeight));
            return y+r.lines.length*r.lineHeight;
        }

        // ─── Date capsule, font size passed in ────────────────────
        function tDate(ctx, c, cx, y, maxW, fs) {
            const ds=`${c.s||'—'}  —  ${c.e||'—'}`;
            ctx.font=`normal ${fs}px Cairo`;
            const dW=Math.min(ctx.measureText(ds).width+fs*1.7,maxW);
            const dH=fs*1.55, dX=cx-dW/2;
            ctx.fillStyle='rgba(42,99,100,0.08)'; drawRoundedRect(ctx,dX,y,dW,dH,dH/2);
            ctx.strokeStyle=TC.goldL; ctx.lineWidth=3; ctx.strokeRect(dX,y,dW,dH);
            ctx.fillStyle=TC.sub; ctx.textAlign='center'; ctx.fillText(ds,cx,y+dH*0.67);
            return y+dH;
        }

        function tCardH(startY, n, gap) {
            const natural = Math.floor((T_H-startY-240-gap*(n-1))/n);
            return Math.max(900, Math.min(1650, natural));
        }

        // ==========================================================
        // TEMPLATE 12 — الشريط الجانبي الفاخر
        // شريط أخضر يمين يحمل رقم ذهبي ضخم + اسم المنسق
        // جميع العناصر نسبية من ارتفاع البطاقة
        // ==========================================================
        function drawReferenceOneScreenCard(ctx, list, stats, pn, tp, si, logo) {
            tBg(ctx);
            let y=tHdr(ctx,logo,pn,tp)+55;
            y=tSts(ctx,stats,y)+65;
            y=tIntr(ctx,y)+30;
            const n=list.length||1, gap=50;
            const cH=tCardH(y,n,gap);
            for(let i=0;i<n;i++){if(list[i])t12(ctx,list[i],T_MX,y,T_W-T_MX*2,cH,si+i+1);y+=cH+gap;}
            tFtr(ctx);
        }

        function t12(ctx,c,x,y,w,h,idx){
            const sW=500;
            // card base
            tSh(ctx,52,18); ctx.fillStyle=TC.bgCard; drawRoundedRect(ctx,x,y,w,h,30); tCl(ctx);
            ctx.strokeStyle=TC.line; ctx.lineWidth=3; ctx.strokeRect(x,y,w,h);
            // green right sidebar (clip)
            ctx.save();
            ctx.beginPath(); ctx.moveTo(x+w-sW,y); ctx.lineTo(x+w-30,y);
            ctx.arcTo(x+w,y,x+w,y+30,30); ctx.lineTo(x+w,y+h-30);
            ctx.arcTo(x+w,y+h,x+w-30,y+h,30); ctx.lineTo(x+w-sW,y+h); ctx.closePath(); ctx.clip();
            const gS=ctx.createLinearGradient(x+w-sW,y,x+w,y+h);
            gS.addColorStop(0,TC.green); gS.addColorStop(1,TC.greenD);
            ctx.fillStyle=gS; ctx.fillRect(x+w-sW,y,sW,h); ctx.restore();
            ctx.fillStyle=TC.gold; ctx.fillRect(x+w-sW,y+25,9,h-50);
            // sidebar: index number (proportional to h)
            const numFs=Math.min(280,h*0.20);
            ctx.fillStyle=TC.gold; ctx.textAlign='center'; ctx.font=`bold ${numFs}px Cairo`;
            ctx.fillText(String(idx),x+w-sW/2,y+h*0.27);
            // separator
            ctx.strokeStyle='rgba(199,176,140,0.55)'; ctx.lineWidth=4;
            ctx.beginPath(); ctx.moveTo(x+w-sW+75,y+h*0.33); ctx.lineTo(x+w-75,y+h*0.33); ctx.stroke();
            // coordinator label
            const coLblFs=Math.min(75,h*0.055);
            ctx.fillStyle='rgba(249,249,249,0.60)'; ctx.font=`normal ${coLblFs}px Cairo`;
            ctx.fillText('منسق التدريب',x+w-sW/2,y+h*0.55);
            ctx.strokeStyle='rgba(199,176,140,0.45)'; ctx.lineWidth=2.5;
            ctx.beginPath(); ctx.moveTo(x+w-sW+80,y+h*0.58); ctx.lineTo(x+w-80,y+h*0.58); ctx.stroke();
            // coordinator name (LARGE)
            const coNameFs=Math.min(130,h*0.09);
            ctx.fillStyle='#FFFFFF'; ctx.font=`bold ${coNameFs}px Cairo`;
            const spLines=wrapTextSimple(ctx,c.sp||'—',sW-70).slice(0,2);
            spLines.forEach((l,li)=>ctx.fillText(l,x+w-sW/2,y+h*0.67+li*(coNameFs*1.35)));
            // --- body ---
            const bW=w-sW-44, midX=x+bW/2;
            // chip row — 35% of card height, anchored at bottom
            const chipH=Math.min(500,h*0.34);
            const chipY=y+h-chipH-44;
            tChip(ctx,c,x+44,chipY,bW-80,chipH);
            // title + date vertically centered in space above chips
            const availH=chipY-y;
            const titleFs=Math.min(190,h*0.13);
            const dateFs=Math.min(95,h*0.068);
            const lh=titleFs*1.40;
            const tRes=wrapTextSmart(ctx,c.n,bW-120,2,titleFs);
            const titleBlockH=tRes.lines.length*lh+(dateFs*1.55)+50;
            const startY=y+(availH-titleBlockH)/2+titleFs;
            ctx.fillStyle=TC.ink; ctx.textAlign='center';
            ctx.font=`bold ${tRes.fontSize}px Cairo`;
            let ty=startY;
            tRes.lines.forEach(l=>{ctx.fillText(l,midX,ty);ty+=lh;});
            ty+=30;
            tDate(ctx,c,midX,ty,bW-140,dateFs);
        }

        // ==========================================================
        // TEMPLATE 13 — الخط الزمني العربي
        // خط زمني رأسي يمين، دوائر مرقمة، بطاقات تمتد يساراً
        // ==========================================================
        function drawReferenceTwoScreenCard(ctx, list, stats, pn, tp, si, logo) {
            tBg(ctx);
            let y=tHdr(ctx,logo,pn,tp)+55;
            y=tSts(ctx,stats,y)+65;
            y=tIntr(ctx,y)+30;
            const n=list.length||1, gap=56;
            const tlX=T_W-T_MX-270;
            const cW=tlX-T_MX-72;
            const cH=tCardH(y,n,gap);
            ctx.strokeStyle='rgba(199,176,140,0.62)'; ctx.lineWidth=8; ctx.setLineDash([46,30]);
            ctx.beginPath(); ctx.moveTo(tlX,y-35); ctx.lineTo(tlX,T_H-248); ctx.stroke(); ctx.setLineDash([]);
            for(let i=0;i<n;i++){
                if(list[i]) t13(ctx,list[i],T_MX,y+i*(cH+gap),cW,cH,si+i+1,tlX);
            }
            tFtr(ctx);
        }

        function t13(ctx,c,x,y,w,h,idx,tlX){
            const mY=y+h/2;
            const nodeR=Math.min(92,h*0.07);
            tSh(ctx,24,9,'rgba(26,68,69,0.22)');
            ctx.fillStyle=TC.green; ctx.beginPath(); ctx.arc(tlX,mY,nodeR,0,Math.PI*2); ctx.fill(); tCl(ctx);
            ctx.strokeStyle=TC.gold; ctx.lineWidth=9; ctx.beginPath(); ctx.arc(tlX,mY,nodeR,0,Math.PI*2); ctx.stroke();
            ctx.fillStyle=TC.gold; ctx.textAlign='center'; ctx.font=`bold ${nodeR*1.10}px Cairo`;
            ctx.fillText(String(idx),tlX,mY+nodeR*0.38);
            ctx.strokeStyle=TC.gold; ctx.lineWidth=5; ctx.setLineDash([22,15]);
            ctx.beginPath(); ctx.moveTo(tlX-nodeR-2,mY); ctx.lineTo(x+w,mY); ctx.stroke(); ctx.setLineDash([]);
            // card
            tSh(ctx,48,15); ctx.fillStyle=TC.bgCard; drawRoundedRect(ctx,x,y,w,h,28); tCl(ctx);
            ctx.strokeStyle=TC.line; ctx.lineWidth=3; ctx.strokeRect(x,y,w,h);
            // date strip top
            const dSH=Math.min(130,h*0.11);
            ctx.save();
            ctx.beginPath(); ctx.moveTo(x+28,y); ctx.lineTo(x+w-28,y);
            ctx.arcTo(x+w,y,x+w,y+28,28); ctx.lineTo(x+w,y+dSH);
            ctx.lineTo(x,y+dSH); ctx.lineTo(x,y+28); ctx.arcTo(x,y,x+28,y,28); ctx.closePath(); ctx.clip();
            const gD=ctx.createLinearGradient(x,y,x+w,y); gD.addColorStop(0,TC.greenD); gD.addColorStop(1,TC.green);
            ctx.fillStyle=gD; ctx.fillRect(x,y,w,dSH); ctx.restore();
            ctx.fillStyle=TC.gold; ctx.fillRect(x,y+dSH-6,w,6);
            const dateInFs=Math.min(90,dSH*0.60);
            ctx.fillStyle=TC.gold; ctx.textAlign='center'; ctx.font=`bold ${dateInFs}px Cairo`;
            ctx.fillText(`${c.s||'—'}  —  ${c.e||'—'}`,x+w/2,y+dSH*0.70);
            // coordinator strip bottom
            const coSH=Math.min(130,h*0.11);
            ctx.fillStyle='rgba(42,99,100,0.07)'; drawRoundedRect(ctx,x+40,y+h-coSH,w-80,coSH,0);
            ctx.fillStyle=TC.green; drawRoundedRect(ctx,x+40,y+h-coSH,w-80,coSH,22);
            ctx.fillStyle=TC.gold; ctx.fillRect(x+40,y+h-coSH,w-80,6);
            const coFs=Math.min(100,coSH*0.60);
            ctx.fillStyle='#FFFFFF'; ctx.textAlign='center'; ctx.font=`bold ${coFs}px Cairo`;
            ctx.fillText(`منسق التدريب:  ${c.sp||'—'}`,x+w/2,y+h-coSH*0.34);
            // chip row (proportional, above coordinator)
            const chipH=Math.min(480,h*0.34);
            const chipY=y+h-coSH-chipH-20;
            tChip(ctx,c,x+40,chipY,w-80,chipH);
            // title centered in remaining space
            const bodyH=chipY-y-dSH-20;
            const titleFs=Math.min(180,h*0.12);
            const tRes=wrapTextSmart(ctx,c.n,w-100,2,titleFs);
            const titleBlockH=tRes.lines.length*titleFs*1.40;
            const titleY=y+dSH+20+(bodyH-titleBlockH)/2+titleFs;
            ctx.fillStyle=TC.ink; ctx.textAlign='center'; ctx.font=`bold ${tRes.fontSize}px Cairo`;
            tRes.lines.forEach((l,i)=>ctx.fillText(l,x+w/2,titleY+i*tRes.lineHeight));
        }

        // ==========================================================
        // TEMPLATE 14 — البطاقات التنفيذية الكبيرة
        // رقم watermark ضخم + لوحة منسق يمين + زخارف هندسية
        // ==========================================================
        function drawReferenceThreeScreenCard(ctx, list, stats, pn, tp, si, logo) {
            tBg(ctx);
            let y=tHdr(ctx,logo,pn,tp)+55;
            y=tSts(ctx,stats,y)+65;
            y=tIntr(ctx,y)+30;
            const n=list.length||1, gap=50;
            const cH=tCardH(y,n,gap);
            for(let i=0;i<n;i++){if(list[i])t14(ctx,list[i],T_MX,y,T_W-T_MX*2,cH,si+i+1);y+=cH+gap;}
            tFtr(ctx);
        }

        function t14(ctx,c,x,y,w,h,idx){
            const cpW=490;
            tSh(ctx,52,18); ctx.fillStyle=TC.bgCard; drawRoundedRect(ctx,x,y,w,h,30); tCl(ctx);
            // geometric bg
            ctx.save(); ctx.globalAlpha=0.033; ctx.fillStyle=TC.green;
            ctx.beginPath(); ctx.arc(x+w,y+h,h*0.68,Math.PI,Math.PI*1.5,true); ctx.fill();
            ctx.beginPath(); ctx.arc(x,y,190,0,Math.PI*0.5); ctx.fill();
            ctx.globalAlpha=1; ctx.restore();
            // watermark number
            const wmFs=Math.min(600,h*0.44);
            ctx.save(); ctx.globalAlpha=0.052; ctx.fillStyle=TC.gold;
            ctx.textAlign='left'; ctx.font=`bold ${wmFs}px Cairo`;
            ctx.fillText(String(idx),x+55,y+h-55); ctx.restore();
            ctx.strokeStyle=TC.line; ctx.lineWidth=3; ctx.strokeRect(x,y,w,h);
            // right coordinator panel
            ctx.save();
            ctx.beginPath(); ctx.moveTo(x+w-cpW,y); ctx.lineTo(x+w-30,y);
            ctx.arcTo(x+w,y,x+w,y+30,30); ctx.lineTo(x+w,y+h-30);
            ctx.arcTo(x+w,y+h,x+w-30,y+h,30); ctx.lineTo(x+w-cpW,y+h); ctx.closePath(); ctx.clip();
            const gP=ctx.createLinearGradient(x+w-cpW,y,x+w,y+h);
            gP.addColorStop(0,TC.green); gP.addColorStop(1,TC.greenD);
            ctx.fillStyle=gP; ctx.fillRect(x+w-cpW,y,cpW,h); ctx.restore();
            ctx.fillStyle=TC.gold; ctx.fillRect(x+w-cpW,y+22,9,h-44);
            // panel: index
            const pNumFs=Math.min(240,h*0.17);
            ctx.fillStyle=TC.gold; ctx.textAlign='center'; ctx.font=`bold ${pNumFs}px Cairo`;
            ctx.fillText(String(idx),x+w-cpW/2,y+h*0.22);
            ctx.strokeStyle='rgba(199,176,140,0.55)'; ctx.lineWidth=4;
            ctx.beginPath(); ctx.moveTo(x+w-cpW+65,y+h*0.26); ctx.lineTo(x+w-65,y+h*0.26); ctx.stroke();
            // person icon
            ctx.save(); ctx.strokeStyle='rgba(199,176,140,0.80)'; ctx.lineWidth=7; ctx.lineCap='round';
            const iX=x+w-cpW/2, iY=y+h*0.50;
            ctx.beginPath(); ctx.arc(iX,iY-h*0.08,h*0.045,0,Math.PI*2); ctx.stroke();
            ctx.beginPath(); ctx.arc(iX,iY+h*0.02,h*0.082,Math.PI,Math.PI*2,true); ctx.stroke(); ctx.restore();
            // coordinator text
            const coLblFs=Math.min(74,h*0.054);
            ctx.fillStyle='rgba(249,249,249,0.58)'; ctx.font=`normal ${coLblFs}px Cairo`;
            ctx.fillText('منسق التدريب',x+w-cpW/2,y+h*0.67);
            ctx.strokeStyle='rgba(199,176,140,0.44)'; ctx.lineWidth=2.5;
            ctx.beginPath(); ctx.moveTo(x+w-cpW+70,y+h*0.70); ctx.lineTo(x+w-70,y+h*0.70); ctx.stroke();
            const coNameFs=Math.min(128,h*0.090);
            ctx.fillStyle='#FFFFFF'; ctx.font=`bold ${coNameFs}px Cairo`;
            wrapTextSimple(ctx,c.sp||'—',cpW-58).slice(0,2).forEach((l,li)=>
                ctx.fillText(l,x+w-cpW/2,y+h*0.78+li*(coNameFs*1.38)));
            // body
            const bW=w-cpW-44, midX=x+bW/2;
            const chipH=Math.min(500,h*0.34);
            const chipY=y+h-chipH-44;
            tChip(ctx,c,x+44,chipY,bW-85,chipH);
            const availH=chipY-y;
            const titleFs=Math.min(190,h*0.13);
            const dateFs=Math.min(95,h*0.068);
            const tRes=wrapTextSmart(ctx,c.n,bW-130,2,titleFs);
            const lh=tRes.fontSize*1.40;
            const titleBlockH=tRes.lines.length*lh+(dateFs*1.55)+50;
            let ty=y+(availH-titleBlockH)/2+tRes.fontSize;
            ctx.fillStyle=TC.ink; ctx.textAlign='center'; ctx.font=`bold ${tRes.fontSize}px Cairo`;
            tRes.lines.forEach(l=>{ctx.fillText(l,midX,ty);ty+=lh;});
            ty+=30;
            tDate(ctx,c,midX,ty,bW-150,dateFs);
        }

        // ==========================================================
        // TEMPLATE 15 — الشبكة الثنائية 2×2
        // عمودان وصفان، كل بطاقة مكثفة بمعلومات واضحة
        // ==========================================================
        function drawReferenceFourScreenCard(ctx, list, stats, pn, tp, si, logo) {
            tBg(ctx);
            let y=tHdr(ctx,logo,pn,tp)+55;
            y=tSts(ctx,stats,y)+65;
            y=tIntr(ctx,y)+30;
            const cols=2, gap=75, mX=T_MX;
            const cW=(T_W-mX*2-gap)/cols;
            const rows=Math.ceil(Math.min(list.length,4)/cols);
            const cH=Math.max(900,Math.floor((T_H-y-240-gap*(rows-1))/rows));
            for(let i=0;i<Math.min(list.length,4);i++){
                if(!list[i]) continue;
                const col=i%cols, row=Math.floor(i/cols);
                t15(ctx,list[i],mX+col*(cW+gap),y+row*(cH+gap),cW,cH,si+i+1);
            }
            tFtr(ctx);
        }

        function t15(ctx,c,x,y,w,h,idx){
            const topH=Math.min(280,h*0.17);
            const botH=Math.min(175,h*0.11);
            // card
            tSh(ctx,46,15); ctx.fillStyle=TC.bgCard; drawRoundedRect(ctx,x,y,w,h,28); tCl(ctx);
            ctx.strokeStyle=TC.line; ctx.lineWidth=3; ctx.strokeRect(x,y,w,h);
            // top green strip
            ctx.save();
            ctx.beginPath(); ctx.moveTo(x+28,y); ctx.lineTo(x+w-28,y);
            ctx.arcTo(x+w,y,x+w,y+28,28); ctx.lineTo(x+w,y+topH);
            ctx.lineTo(x,y+topH); ctx.lineTo(x,y+28); ctx.arcTo(x,y,x+28,y,28); ctx.closePath(); ctx.clip();
            const gT=ctx.createLinearGradient(x,y,x+w,y); gT.addColorStop(0,TC.greenD); gT.addColorStop(1,TC.green);
            ctx.fillStyle=gT; ctx.fillRect(x,y,w,topH); ctx.restore();
            ctx.fillStyle=TC.gold; ctx.fillRect(x,y+topH-7,w,7);
            // number (gold, left side of strip)
            const numFs=Math.min(200,topH*0.85);
            ctx.fillStyle=TC.gold; ctx.textAlign='left'; ctx.font=`bold ${numFs}px Cairo`;
            ctx.fillText(String(idx),x+50,y+topH-12);
            // title in strip (white, centered, proportional)
            ctx.fillStyle='#FFFFFF'; ctx.textAlign='center';
            const tRes15=wrapTextSmart(ctx,c.n,w-260,2,Math.min(90,topH*0.42));
            ctx.font=`bold ${tRes15.fontSize}px Cairo`;
            let ty15=y+(topH-tRes15.lines.length*tRes15.lineHeight)/2+tRes15.fontSize*0.88;
            tRes15.lines.forEach(l=>{ctx.fillText(l,x+w/2+40,ty15);ty15+=tRes15.lineHeight;});
            // date row
            const dtFs=Math.min(85,h*0.055);
            const dtY=y+topH+h*0.05;
            ctx.fillStyle=TC.sub; ctx.textAlign='center'; ctx.font=`normal ${dtFs}px Cairo`;
            ctx.fillText(`${c.s||'—'}  —  ${c.e||'—'}`,x+w/2,dtY);
            // chips (2 rows fitting available space)
            const chipArea=h-topH-botH-h*0.05-dtFs*1.5-20;
            const chipH=Math.min(220,Math.floor(chipArea/2)-20);
            const chipGap=20;
            const chipY1=dtY+dtFs*0.6+h*0.03;
            const chipY2=chipY1+chipH+chipGap;
            const r1=[['مكان التنفيذ',c.loc,0],['الفترة',c.p,1],['الحالة',c.st,2]];
            const r2=[['القاعة',c.r,3],['الطابق',shouldHideFloorValue(c.f)?'—':c.f,4]];
            const bW1=(w-80-chipGap*2)/3, bW2=(w-80-chipGap)/2;
            const vFs=Math.min(130,chipH*0.36), lFs=Math.min(55,chipH*0.18);
            const dR=(arr,bWW,ry)=>arr.forEach((it,i)=>{
                const bx=x+40+i*(bWW+chipGap);
                ctx.fillStyle=TC.soft; drawRoundedRect(ctx,bx,ry,bWW,chipH,20);
                ctx.strokeStyle=TC.goldL; ctx.lineWidth=2.5; ctx.strokeRect(bx,ry,bWW,chipH);
                ctx.fillStyle=TC.gold; ctx.fillRect(bx,ry,bWW,6);
                tChipIcon(ctx,it[2],bx+bWW/2,ry+chipH*0.28,chipH*0.18);
                ctx.fillStyle=TC.ink; ctx.textAlign='center'; ctx.font=`bold ${vFs}px Cairo`;
                let v=String(it[1]||'—'); while(ctx.measureText(v).width>bWW-30&&v.length>0)v=v.slice(0,-1);
                if(v!==String(it[1]||'—'))v+='...';
                ctx.fillText(v,bx+bWW/2,ry+chipH*0.64);
                ctx.fillStyle=TC.muted; ctx.font=`normal ${lFs}px Cairo`;
                ctx.fillText(it[0],bx+bWW/2,ry+chipH*0.86);
            });
            dR(r1,bW1,chipY1); dR(r2,bW2,chipY2);
            // bottom coordinator
            const gB=ctx.createLinearGradient(x,y+h-botH,x+w,y+h-botH);
            gB.addColorStop(0,TC.greenD); gB.addColorStop(1,TC.green);
            ctx.fillStyle=gB; drawRoundedRect(ctx,x,y+h-botH,w,botH,28);
            ctx.fillStyle=gB; ctx.fillRect(x,y+h-botH,w,botH-28);
            ctx.fillStyle=TC.gold; ctx.fillRect(x,y+h-botH,w,7);
            const coFs15=Math.min(90,botH*0.55);
            ctx.fillStyle='#FFFFFF'; ctx.textAlign='center'; ctx.font=`bold ${coFs15}px Cairo`;
            ctx.fillText(`منسق التدريب:  ${c.sp||'—'}`,x+w/2,y+h-botH/2+coFs15*0.35);
        }

        // ==========================================================
        // TEMPLATE 16 — لوحة العمليات الفاخرة
        // صفوف أفقية: شارة يمين + محتوى وسط + منسق يسار
        // ==========================================================
        function drawReferenceFiveScreenCard(ctx, list, stats, pn, tp, si, logo) {
            tBg(ctx);
            let y=tHdr(ctx,logo,pn,tp)+55;
            y=tSts(ctx,stats,y)+65;
            y=tIntr(ctx,y)+30;
            const n=list.length||1, gap=44;
            const cH=tCardH(y,n,gap);
            for(let i=0;i<n;i++){if(list[i])t16(ctx,list[i],T_MX,y,T_W-T_MX*2,cH,si+i+1);y+=cH+gap;}
            tFtr(ctx);
        }

        function t16(ctx,c,x,y,w,h,idx){
            const bR=Math.min(110,h*0.38);
            const bZW=bR*2+70;
            const coW=Math.min(500,w*0.21);
            const iP=38;
            // card
            tSh(ctx,46,14); ctx.fillStyle=TC.bgCard; drawRoundedRect(ctx,x,y,w,h,28); tCl(ctx);
            ctx.fillStyle=TC.gold; ctx.fillRect(x+w-12,y+24,12,h-48);
            ctx.strokeStyle=TC.line; ctx.lineWidth=3; ctx.strokeRect(x,y,w,h);
            // badge right
            const bdX=x+w-bZW/2-20, bdY=y+h/2;
            tSh(ctx,22,9,'rgba(26,68,69,0.22)');
            ctx.fillStyle=TC.green; ctx.beginPath(); ctx.arc(bdX,bdY,bR,0,Math.PI*2); ctx.fill(); tCl(ctx);
            ctx.strokeStyle=TC.gold; ctx.lineWidth=10; ctx.beginPath(); ctx.arc(bdX,bdY,bR,0,Math.PI*2); ctx.stroke();
            const badgeFs=Math.min(120,bR*1.07);
            ctx.fillStyle=TC.gold; ctx.textAlign='center'; ctx.font=`bold ${badgeFs}px Cairo`;
            ctx.fillText(String(idx),bdX,bdY+badgeFs*0.37);
            // coordinator panel (LEFT)
            const coX=x+iP, coY=y+iP, coH=h-iP*2;
            tSh(ctx,18,6,'rgba(26,68,69,0.11)');
            ctx.fillStyle=TC.green; drawRoundedRect(ctx,coX,coY,coW,coH,22); tCl(ctx);
            ctx.fillStyle=TC.gold; ctx.fillRect(coX+coW-8,coY+14,8,coH-28);
            const coLblFs16=Math.min(72,h*0.052);
            ctx.fillStyle='rgba(249,249,249,0.55)'; ctx.textAlign='center'; ctx.font=`normal ${coLblFs16}px Cairo`;
            ctx.fillText('منسق التدريب',coX+coW/2,coY+coH*0.40);
            ctx.strokeStyle='rgba(199,176,140,0.52)'; ctx.lineWidth=3;
            ctx.beginPath(); ctx.moveTo(coX+44,coY+coH*0.46); ctx.lineTo(coX+coW-44,coY+coH*0.46); ctx.stroke();
            const coNameFs16=Math.min(120,h*0.088);
            ctx.fillStyle='#FFFFFF'; ctx.font=`bold ${coNameFs16}px Cairo`;
            wrapTextSimple(ctx,c.sp||'—',coW-48).slice(0,2).forEach((l,li)=>
                ctx.fillText(l,coX+coW/2,coY+coH*0.57+li*(coNameFs16*1.38)));
            // content zone
            const cntX=coX+coW+iP, cntW=w-coW-bZW-iP*3-28, midX=cntX+cntW/2;
            // chip row anchored to bottom of card
            const chipH16=Math.min(480,h*0.34);
            const chipY16=y+h-chipH16-38;
            tChip(ctx,c,cntX,chipY16,cntW,chipH16);
            // title + date centered above chips
            const availH16=chipY16-y;
            const titleFs16=Math.min(185,h*0.13);
            const dateFs16=Math.min(90,h*0.065);
            const tRes16=wrapTextSmart(ctx,c.n,cntW-60,2,titleFs16);
            const lh16=tRes16.fontSize*1.40;
            const titleBlockH16=tRes16.lines.length*lh16+(dateFs16*1.55)+50;
            let ty16=y+(availH16-titleBlockH16)/2+tRes16.fontSize;
            ctx.fillStyle=TC.ink; ctx.textAlign='center'; ctx.font=`bold ${tRes16.fontSize}px Cairo`;
            tRes16.lines.forEach(l=>{ctx.fillText(l,midX,ty16);ty16+=lh16;});
            ty16+=30;
            tDate(ctx,c,midX,ty16,cntW-60,dateFs16);
        }

        // ==========================================================
        // TEMPLATES 17-21 — مبنية على نفس منهج iOS HD
        // المبدأ: احسب مجموع المحتوى أولاً ثم مركزه في البطاقة
        // كل مقياس نسبة من h (ارتفاع البطاقة) لا أرقام ثابتة
        // ==========================================================

        // ── مساعد: رسم chips بمقاسات نسبية من h ──────────────────
        function nChips(ctx, c, x, y, w, h) {
            // RTL order: الطابق(right) → القاعة → الحالة → الفترة → مكان التنفيذ(left)
            const items = [
                ['مكان التنفيذ', c.loc,    0],
                ['الفترة',        c.p,     1],
                ['الحالة',        c.st,    2],
                ['القاعة',        c.r,     3],
                ['الطابق',        shouldHideFloorValue(c.f)?'—':c.f, 4],
            ];
            const gap = w * 0.012, bW = (w - gap*4) / 5;
            const icS = h * 0.22, vFs = Math.min(h*0.36, 180), lFs = Math.min(h*0.17, 76);
            items.forEach((it, i) => {
                const bx = x + i*(bW+gap);
                ctx.fillStyle = TC.soft; drawRoundedRect(ctx,bx,y,bW,h,20);
                ctx.strokeStyle = TC.goldL; ctx.lineWidth = 3; ctx.strokeRect(bx,y,bW,h);
                ctx.fillStyle = TC.gold; ctx.fillRect(bx,y,bW,7);
                tChipIcon(ctx, it[2], bx+bW/2, y+h*0.30, icS);
                // القيمة: قلّص الخط ليتسع النص كاملاً قبل اللجوء للقص
                ctx.fillStyle = TC.ink; ctx.textAlign = 'center';
                const v = String(it[1]||'—'); let vf = vFs;
                ctx.font = `bold ${vf}px Cairo`;
                while(ctx.measureText(v).width > bW-28 && vf > 44){ vf -= 4; ctx.font = `bold ${vf}px Cairo`; }
                let vt = v;
                while(ctx.measureText(vt).width > bW-28 && vt.length > 0) vt = vt.slice(0,-1);
                if(vt !== v) vt += '…';
                ctx.fillText(vt, bx+bW/2, y+h*0.65);
                ctx.fillStyle = TC.muted; ctx.font = `normal ${lFs}px Cairo`;
                ctx.fillText(it[0], bx+bW/2, y+h*0.87);
            });
        }

        // ── مساعد: رسم chips صفّين (3+2) بمقاسات نسبية من h ─────
        function nChips32(ctx, c, x, y, w, h) {
            const row1 = [['الطابق',shouldHideFloorValue(c.f)?'—':c.f,4],['القاعة',c.r,3],['الحالة',c.st,2]];
            const row2 = [['الفترة',c.p,1],['مكان التنفيذ',c.loc,0]];
            const gap = w * 0.012, cGap = h * 0.045;
            const rH = (h - cGap) / 2;
            const bW1 = (w - gap*2) / 3, bW2 = (w - gap) / 2;
            const icS = rH * 0.22, vFs = Math.min(rH*0.36, 170), lFs = Math.min(rH*0.17, 72);
            const drawRow = (arr, bWW, ry) => arr.forEach((it, i) => {
                const bx = x + i*(bWW+gap);
                ctx.fillStyle = TC.soft; drawRoundedRect(ctx,bx,ry,bWW,rH,18);
                ctx.strokeStyle = TC.goldL; ctx.lineWidth = 3; ctx.strokeRect(bx,ry,bWW,rH);
                ctx.fillStyle = TC.gold; ctx.fillRect(bx,ry,bWW,7);
                tChipIcon(ctx, it[2], bx+bWW/2, ry+rH*0.30, icS);
                ctx.fillStyle = TC.ink; ctx.textAlign = 'center';
                const v = String(it[1]||'—'); let vf = vFs;
                ctx.font = `bold ${vf}px Cairo`;
                while(ctx.measureText(v).width > bWW-24 && vf > 40){ vf -= 4; ctx.font = `bold ${vf}px Cairo`; }
                let vt = v;
                while(ctx.measureText(vt).width > bWW-24 && vt.length > 0) vt = vt.slice(0,-1);
                if(vt !== v) vt += '…';
                ctx.fillText(vt, bx+bWW/2, ry+rH*0.65);
                ctx.fillStyle = TC.muted; ctx.font = `normal ${lFs}px Cairo`;
                ctx.fillText(it[0], bx+bWW/2, ry+rH*0.87);
            });
            drawRow(row1, bW1, y);
            drawRow(row2, bW2, y+rH+cGap);
        }

        // ── مساعد: تاريخ مقيّد بالعرض + أيقونة تقويم ذهبية ──────
        // يصغّر الخط تلقائياً حتى يتسع داخل maxW (لا يطفح أبداً)
        function nDate(ctx, c, cx, midY, maxW, maxFs) {
            const ds = `${c.s||'—'}   ←   ${c.e||'—'}`;
            let fs = Math.round(maxFs);
            ctx.font = `bold ${fs}px Cairo`;
            while (ctx.measureText(ds).width > maxW && fs > 22) { fs -= 2; ctx.font = `bold ${fs}px Cairo`; }
            const tw = ctx.measureText(ds).width;
            const icoS = fs*0.62, icoGap = fs*0.55;
            const groupW = tw + icoGap + icoS;
            // RTL: التقويم على اليمين، النص على يساره
            const icoCx = cx + groupW/2 - icoS/2;
            const txtCx = cx - (icoGap + icoS)/2;
            // calendar glyph
            ctx.save();
            ctx.strokeStyle = TC.gold; ctx.lineWidth = Math.max(3, fs*0.07); ctx.lineJoin='round';
            const r = icoS/2;
            drawRoundedRect2(ctx, icoCx-r, midY-r*0.82, r*2, r*1.7, r*0.22, true);
            ctx.beginPath(); ctx.moveTo(icoCx-r, midY-r*0.30); ctx.lineTo(icoCx+r, midY-r*0.30); ctx.stroke();
            ctx.beginPath(); ctx.moveTo(icoCx-r*0.45, midY-r*0.82-r*0.30); ctx.lineTo(icoCx-r*0.45, midY-r*0.82+r*0.10);
            ctx.moveTo(icoCx+r*0.45, midY-r*0.82-r*0.30); ctx.lineTo(icoCx+r*0.45, midY-r*0.82+r*0.10); ctx.stroke();
            ctx.restore();
            // date text
            ctx.fillStyle = TC.sub; ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
            ctx.font = `bold ${fs}px Cairo`;
            ctx.fillText(ds, txtCx, midY);
            ctx.textBaseline = 'alphabetic';
            return fs;
        }
        // stroke-able rounded rect helper (للأيقونات)
        function drawRoundedRect2(ctx, x, y, w, h, r, stroke) {
            ctx.beginPath();
            ctx.moveTo(x+r,y); ctx.lineTo(x+w-r,y); ctx.quadraticCurveTo(x+w,y,x+w,y+r);
            ctx.lineTo(x+w,y+h-r); ctx.quadraticCurveTo(x+w,y+h,x+w-r,y+h);
            ctx.lineTo(x+r,y+h); ctx.quadraticCurveTo(x,y+h,x,y+h-r);
            ctx.lineTo(x,y+r); ctx.quadraticCurveTo(x,y,x+r,y); ctx.closePath();
            if(stroke) ctx.stroke(); else ctx.fill();
        }

        // ── Template 17: شبكة 2×2 المتميزة ──────────────────────
        function drawRef17ScreenCard(ctx, list, stats, pn, tp, si, logo) {
            tBg(ctx);
            let y = tHdr(ctx, logo, pn, tp) + 55;
            y = tSts(ctx, stats, y) + 65;
            y = tIntr(ctx, y) + 30;
            const cols=2, gap=70, mX=T_MX;
            const cW = (T_W - mX*2 - gap) / cols;
            const rows = 2;
            const cH = Math.max(800, Math.floor((T_H - y - 240 - gap) / rows));
            for(let i=0; i<Math.min(list.length,4); i++) {
                if(!list[i]) continue;
                const col=i%cols, row=Math.floor(i/cols);
                t17(ctx, list[i], mX+col*(cW+gap), y+row*(cH+gap), cW, cH, si+i+1);
            }
            tFtr(ctx);
        }

        function t17(ctx, c, x, y, w, h, idx) {
            // ── المناطق الثابتة ──────────────────────────────────
            const HDR = Math.round(h * 0.175);   // رأس أخضر
            const BOT = Math.round(h * 0.090);   // شريط منسق
            const PAD = Math.round(h * 0.018);   // padding داخلي

            // ── البطاقة الأساسية ──────────────────────────────────
            tSh(ctx,48,14); ctx.fillStyle=TC.bgCard; drawRoundedRect(ctx,x,y,w,h,26); tCl(ctx);
            ctx.strokeStyle=TC.line; ctx.lineWidth=3; ctx.strokeRect(x,y,w,h);

            // ── الرأس الأخضر ──────────────────────────────────────
            ctx.save();
            ctx.beginPath(); ctx.moveTo(x+26,y); ctx.lineTo(x+w-26,y); ctx.arcTo(x+w,y,x+w,y+26,26);
            ctx.lineTo(x+w,y+HDR); ctx.lineTo(x,y+HDR); ctx.lineTo(x,y+26); ctx.arcTo(x,y,x+26,y,26); ctx.closePath(); ctx.clip();
            const gH17=ctx.createLinearGradient(x,y,x+w,y+HDR); gH17.addColorStop(0,TC.greenD); gH17.addColorStop(1,TC.green);
            ctx.fillStyle=gH17; ctx.fillRect(x,y,w,HDR); ctx.restore();
            ctx.fillStyle=TC.gold; ctx.fillRect(x,y+HDR-6,w,6);

            // ── شارة الرقم (يمين الرأس) ───────────────────────────
            const numR = Math.round(HDR * 0.40);
            const numX = x+w-numR-PAD*1.5, numY = y+HDR/2;
            ctx.fillStyle='rgba(199,176,140,0.18)'; ctx.beginPath(); ctx.arc(numX,numY,numR,0,Math.PI*2); ctx.fill();
            ctx.strokeStyle=TC.gold; ctx.lineWidth=5; ctx.beginPath(); ctx.arc(numX,numY,numR,0,Math.PI*2); ctx.stroke();
            ctx.fillStyle=TC.gold; ctx.textAlign='center'; ctx.font=`bold ${Math.round(numR*1.05)}px Cairo`;
            ctx.fillText(String(idx), numX, numY+numR*0.37);

            // ── العنوان في الرأس (يتقلّص ليتسع) ──────────────────
            const tAvailW = w - numR*2.3 - PAD*3;
            const tFs17 = Math.round(HDR * 0.26);
            const tRes17 = wrapTextSmart(ctx, c.n, tAvailW, 2, tFs17);
            ctx.fillStyle='#FFFFFF'; ctx.textAlign='right'; ctx.font=`bold ${tRes17.fontSize}px Cairo`;
            let ty17 = y + (HDR - tRes17.lines.length*tRes17.lineHeight)/2 + tRes17.fontSize*0.88;
            tRes17.lines.forEach(l=>{ ctx.fillText(l, x+tAvailW+PAD*2, ty17); ty17+=tRes17.lineHeight; });

            // ── المنطقة الداخلية: التاريخ ثم chips تملأ الباقي ────
            const innerH = h - HDR - BOT;
            const dtBandH = Math.round(innerH * 0.16);
            const dtTop   = y + HDR;
            nDate(ctx, c, x+w/2, dtTop + dtBandH/2, w - PAD*6, innerH*0.075);
            const chipsTop = dtTop + dtBandH + PAD;
            const chipsH   = (y + h - BOT - PAD) - chipsTop;
            nChips32(ctx, c, x+PAD*2, chipsTop, w-PAD*4, chipsH);

            // ── شريط المنسق (أسفل) ───────────────────────────────
            ctx.save();
            ctx.beginPath(); ctx.moveTo(x,y+h-BOT); ctx.lineTo(x+w,y+h-BOT);
            ctx.lineTo(x+w,y+h-26); ctx.arcTo(x+w,y+h,x+w-26,y+h,26); ctx.lineTo(x+26,y+h); ctx.arcTo(x,y+h,x,y+h-26,26); ctx.closePath(); ctx.clip();
            const gB17=ctx.createLinearGradient(x,y+h-BOT,x+w,y+h); gB17.addColorStop(0,TC.greenD); gB17.addColorStop(1,TC.green);
            ctx.fillStyle=gB17; ctx.fillRect(x,y+h-BOT,w,BOT); ctx.restore();
            ctx.fillStyle=TC.gold; ctx.fillRect(x,y+h-BOT,w,6);
            ctx.fillStyle='#FFFFFF'; ctx.textAlign='center';
            ctx.font=`bold ${Math.round(BOT*0.48)}px Cairo`;
            ctx.fillText(`منسق التدريب:  ${c.sp||'—'}`, x+w/2, y+h-BOT/2+BOT*0.17);
        }

        // ── Template 18: أفقي — شريط أخضر يمين ──────────────────
        function drawRef18ScreenCard(ctx, list, stats, pn, tp, si, logo) {
            tBg(ctx);
            let y = tHdr(ctx, logo, pn, tp) + 55;
            y = tSts(ctx, stats, y) + 65;
            y = tIntr(ctx, y) + 30;
            const n=list.length||1, gap=52;
            const cH=tCardH(y, n, gap);
            for(let i=0;i<n;i++){if(list[i])t18(ctx,list[i],T_MX,y,T_W-T_MX*2,cH,si+i+1);y+=cH+gap;}
            tFtr(ctx);
        }

        function t18(ctx, c, x, y, w, h, idx) {
            const PNL = Math.round(w * 0.195);  // عرض الشريط الأيمن
            const PAD = Math.round(h * 0.04);

            // ── البطاقة الأساسية ──────────────────────────────────
            tSh(ctx,50,16); ctx.fillStyle=TC.bgCard; drawRoundedRect(ctx,x,y,w,h,28); tCl(ctx);
            // رقم watermark خفيف يسار
            ctx.save(); ctx.globalAlpha=0.036; ctx.fillStyle=TC.green;
            ctx.textAlign='left'; ctx.font=`bold ${Math.round(h*0.60)}px Cairo`;
            ctx.fillText(String(idx), x+PAD, y+h*0.82); ctx.restore();
            ctx.strokeStyle=TC.line; ctx.lineWidth=3; ctx.strokeRect(x,y,w,h);

            // ── الشريط الأخضر الأيمن ─────────────────────────────
            ctx.save();
            ctx.beginPath(); ctx.moveTo(x+w-PNL,y); ctx.lineTo(x+w-28,y); ctx.arcTo(x+w,y,x+w,y+28,28);
            ctx.lineTo(x+w,y+h-28); ctx.arcTo(x+w,y+h,x+w-28,y+h,28); ctx.lineTo(x+w-PNL,y+h); ctx.closePath(); ctx.clip();
            const gP=ctx.createLinearGradient(x+w-PNL,y,x+w,y+h); gP.addColorStop(0,TC.green); gP.addColorStop(1,TC.greenD);
            ctx.fillStyle=gP; ctx.fillRect(x+w-PNL,y,PNL,h); ctx.restore();
            ctx.fillStyle=TC.gold; ctx.fillRect(x+w-PNL-6,y+20,6,h-40);
            const pCX=x+w-PNL/2;
            // رقم في الشريط
            ctx.fillStyle=TC.gold; ctx.textAlign='center'; ctx.font=`bold ${Math.round(PNL*0.52)}px Cairo`;
            ctx.fillText(String(idx), pCX, y+h*0.22);
            ctx.strokeStyle='rgba(199,176,140,0.50)'; ctx.lineWidth=4;
            ctx.beginPath(); ctx.moveTo(x+w-PNL+50,y+h*0.28); ctx.lineTo(x+w-50,y+h*0.28); ctx.stroke();
            // أيقونة الشخص
            ctx.save(); ctx.strokeStyle='rgba(199,176,140,0.78)'; ctx.lineWidth=Math.round(h*0.006); ctx.lineCap='round';
            ctx.beginPath(); ctx.arc(pCX,y+h*0.43,h*0.04,0,Math.PI*2); ctx.stroke();
            ctx.beginPath(); ctx.arc(pCX,y+h*0.49,h*0.072,Math.PI,Math.PI*2,true); ctx.stroke(); ctx.restore();
            ctx.fillStyle='rgba(249,249,249,0.55)'; ctx.font=`normal ${Math.round(h*0.044)}px Cairo`;
            ctx.fillText('منسق التدريب', pCX, y+h*0.62);
            ctx.strokeStyle='rgba(199,176,140,0.42)'; ctx.lineWidth=2.5;
            ctx.beginPath(); ctx.moveTo(x+w-PNL+50,y+h*0.65); ctx.lineTo(x+w-50,y+h*0.65); ctx.stroke();
            const cnFs=Math.round(h*0.080);
            ctx.fillStyle='#FFFFFF'; ctx.font=`bold ${cnFs}px Cairo`;
            wrapTextSimple(ctx,c.sp||'—',PNL-50).slice(0,2).forEach((l,li)=>ctx.fillText(l,pCX,y+h*0.74+li*cnFs*1.35));
            ctx.fillStyle=TC.gold; ctx.save(); ctx.translate(pCX,y+h*0.92); ctx.rotate(Math.PI/4); ctx.fillRect(-12,-12,24,24); ctx.restore();

            // ── منطقة المحتوى (يسار الشريط) ─────────────────────
            // المبدأ: احسب المحتوى أولاً ثم مركزه — مثل iOS
            const bW  = w - PNL - PAD;
            const bX  = x + PAD/2;
            const cX  = bX + bW/2;
            const ttlFs = Math.round(h * 0.115);
            const dtFs  = Math.round(h * 0.052);
            const chH   = Math.round(h * 0.310);  // ارتفاع صف الـ chips
            const dtGap = Math.round(h * 0.030);
            const chGap = Math.round(h * 0.035);

            const tRes = wrapTextSmart(ctx, c.n, bW-PAD*2, 2, ttlFs);
            const ttlH = tRes.lines.length * tRes.lineHeight;
            const total = ttlH + dtGap + dtFs*1.6 + chGap + chH;
            let cy = y + (h - total) / 2;

            // عنوان
            ctx.fillStyle=TC.ink; ctx.textAlign='center'; ctx.font=`bold ${tRes.fontSize}px Cairo`;
            tRes.lines.forEach((l,i)=>{ ctx.fillText(l, cX, cy+tRes.fontSize+i*tRes.lineHeight); });
            cy += ttlH + dtGap;
            // تاريخ (مقيّد بالعرض)
            nDate(ctx, c, cX, cy+dtFs*0.7, bW-PAD*2, dtFs);
            cy += Math.round(dtFs*1.6) + chGap;
            // chips (صف واحد 5 عناصر)
            nChips(ctx, c, bX+PAD/2, cy, bW-PAD, chH);
        }

        // ── Template 19: الخط الزمني — يمين RTL ──────────────────
        function drawRef19ScreenCard(ctx, list, stats, pn, tp, si, logo) {
            tBg(ctx);
            let y = tHdr(ctx, logo, pn, tp) + 55;
            y = tSts(ctx, stats, y) + 65;
            y = tIntr(ctx, y) + 30;
            const n=list.length||1, gap=62;
            const tlX=T_W-T_MX-200;
            const cW=tlX-T_MX-90;
            const cH=tCardH(y, n, gap);
            // خط الجدول
            ctx.save(); ctx.strokeStyle=TC.gold; ctx.lineWidth=9; ctx.setLineDash([52,34]);
            ctx.beginPath(); ctx.moveTo(tlX,y-35); ctx.lineTo(tlX,y+n*(cH+gap)-gap+35); ctx.stroke();
            ctx.setLineDash([]); ctx.restore();
            for(let i=0;i<n;i++){if(list[i])t19(ctx,list[i],T_MX,y+i*(cH+gap),cW,cH,si+i+1,tlX);}
            tFtr(ctx);
        }

        function t19(ctx, c, x, y, w, h, idx, tlX) {
            const mY=y+h/2;
            const nodeR=Math.round(h*0.068);
            // نقطة الجدول الزمني
            tSh(ctx,24,10,'rgba(26,68,69,0.22)');
            ctx.fillStyle=TC.green; ctx.beginPath(); ctx.arc(tlX,mY,nodeR,0,Math.PI*2); ctx.fill(); tCl(ctx);
            ctx.strokeStyle=TC.gold; ctx.lineWidth=9; ctx.beginPath(); ctx.arc(tlX,mY,nodeR,0,Math.PI*2); ctx.stroke();
            ctx.fillStyle='#FFFFFF'; ctx.textAlign='center'; ctx.font=`bold ${Math.round(nodeR*1.08)}px Cairo`;
            ctx.fillText(String(idx), tlX, mY+nodeR*0.38);
            ctx.strokeStyle=TC.gold; ctx.lineWidth=5; ctx.setLineDash([22,14]);
            ctx.beginPath(); ctx.moveTo(tlX-nodeR-2,mY); ctx.lineTo(x+w,mY); ctx.stroke(); ctx.setLineDash([]);

            // البطاقة
            tSh(ctx,50,15); ctx.fillStyle=TC.bgCard; drawRoundedRect(ctx,x,y,w,h,28); tCl(ctx);
            ctx.strokeStyle=TC.line; ctx.lineWidth=3; ctx.strokeRect(x,y,w,h);

            // ── شريط التاريخ أعلى ────────────────────────────────
            const DSH = Math.round(h * 0.100);
            ctx.save();
            ctx.beginPath(); ctx.moveTo(x+28,y); ctx.lineTo(x+w-28,y); ctx.arcTo(x+w,y,x+w,y+28,28);
            ctx.lineTo(x+w,y+DSH); ctx.lineTo(x,y+DSH); ctx.lineTo(x,y+28); ctx.arcTo(x,y,x+28,y,28); ctx.closePath(); ctx.clip();
            const gD=ctx.createLinearGradient(x,y,x+w,y+DSH); gD.addColorStop(0,TC.greenD); gD.addColorStop(1,TC.green);
            ctx.fillStyle=gD; ctx.fillRect(x,y,w,DSH); ctx.restore();
            ctx.fillStyle=TC.gold; ctx.fillRect(x,y+DSH-6,w,6);
            ctx.fillStyle=TC.gold; ctx.textAlign='center'; ctx.font=`bold ${Math.round(DSH*0.55)}px Cairo`;
            ctx.fillText(`${c.s||'—'}  ←  ${c.e||'—'}`, x+w/2, y+DSH*0.70);

            // ── شريط المنسق أسفل ──────────────────────────────────
            const CSH = Math.round(h * 0.100);
            ctx.save();
            ctx.beginPath(); ctx.moveTo(x,y+h-CSH); ctx.lineTo(x+w,y+h-CSH);
            ctx.lineTo(x+w,y+h-28); ctx.arcTo(x+w,y+h,x+w-28,y+h,28); ctx.lineTo(x+28,y+h); ctx.arcTo(x,y+h,x,y+h-28,28); ctx.closePath(); ctx.clip();
            ctx.fillStyle=TC.green; ctx.fillRect(x,y+h-CSH,w,CSH); ctx.restore();
            ctx.fillStyle=TC.gold; ctx.fillRect(x,y+h-CSH,w,6);
            ctx.fillStyle='#FFFFFF'; ctx.textAlign='center'; ctx.font=`bold ${Math.round(CSH*0.55)}px Cairo`;
            ctx.fillText(`منسق التدريب:  ${c.sp||'—'}`, x+w/2, y+h-CSH*0.34);

            // ── المنطقة الداخلية: عنوان + chips (iOS method) ──────
            const innerY = y + DSH;
            const innerH = h - DSH - CSH;
            const PAD = Math.round(innerH * 0.04);

            const ttlFs = Math.round(innerH * 0.160);
            const chH   = Math.round(innerH * 0.420);
            const chGap = Math.round(innerH * 0.040);

            const tRes = wrapTextSmart(ctx, c.n, w-PAD*4, 2, ttlFs);
            const ttlH = tRes.lines.length * tRes.lineHeight;
            const total = ttlH + chGap + chH;
            let cy = innerY + (innerH - total) / 2;

            // عنوان
            ctx.fillStyle=TC.ink; ctx.textAlign='center'; ctx.font=`bold ${tRes.fontSize}px Cairo`;
            tRes.lines.forEach((l,i)=>ctx.fillText(l, x+w/2, cy+tRes.fontSize+i*tRes.lineHeight));
            cy += ttlH + chGap;
            // chips
            nChips(ctx, c, x+PAD*2, cy, w-PAD*4, chH);
        }

        // ── Template 20: التنفيذي الفاخر ─────────────────────────
        function drawRef20ScreenCard(ctx, list, stats, pn, tp, si, logo) {
            tBg(ctx);
            let y = tHdr(ctx, logo, pn, tp) + 55;
            y = tSts(ctx, stats, y) + 65;
            y = tIntr(ctx, y) + 30;
            const n=list.length||1, gap=55;
            const cH=tCardH(y, n, gap);
            for(let i=0;i<n;i++){if(list[i])t20(ctx,list[i],T_MX,y,T_W-T_MX*2,cH,si+i+1);y+=cH+gap;}
            tFtr(ctx);
        }

        function t20(ctx, c, x, y, w, h, idx) {
            const PNL = Math.round(w * 0.185);  // لوحة المنسق يمين
            const PAD = Math.round(h * 0.04);

            // ── البطاقة ───────────────────────────────────────────
            tSh(ctx,52,18); ctx.fillStyle=TC.bgCard; drawRoundedRect(ctx,x,y,w,h,30); tCl(ctx);
            // خلفية هندسية خفيفة
            ctx.save(); ctx.globalAlpha=0.022; ctx.fillStyle=TC.green;
            ctx.beginPath(); ctx.arc(x+w,y+h,h*0.68,Math.PI,Math.PI*1.5,true); ctx.fill();
            ctx.beginPath(); ctx.arc(x,y,180,0,Math.PI*0.5); ctx.fill(); ctx.restore();
            ctx.fillStyle=TC.gold; ctx.fillRect(x,y,w,8);
            ctx.strokeStyle=TC.line; ctx.lineWidth=3; ctx.strokeRect(x,y,w,h);

            // رقم watermark يسار (خفيف جداً)
            ctx.save(); ctx.globalAlpha=0.038; ctx.fillStyle=TC.gold;
            ctx.textAlign='left'; ctx.font=`bold ${Math.round(h*0.42)}px Cairo`;
            ctx.fillText(String(idx), x+PAD*2, y+h-PAD*2); ctx.restore();

            // ── لوحة المنسق يمين ──────────────────────────────────
            ctx.save();
            ctx.beginPath(); ctx.moveTo(x+w-PNL,y); ctx.lineTo(x+w-28,y); ctx.arcTo(x+w,y,x+w,y+28,28);
            ctx.lineTo(x+w,y+h-28); ctx.arcTo(x+w,y+h,x+w-28,y+h,28); ctx.lineTo(x+w-PNL,y+h); ctx.closePath(); ctx.clip();
            const gCo=ctx.createLinearGradient(x+w-PNL,y,x+w,y+h); gCo.addColorStop(0,TC.green); gCo.addColorStop(1,TC.greenD);
            ctx.fillStyle=gCo; ctx.fillRect(x+w-PNL,y,PNL,h); ctx.restore();
            ctx.fillStyle=TC.gold; ctx.fillRect(x+w-PNL-6,y+22,6,h-44);
            const pCX=x+w-PNL/2;
            ctx.fillStyle=TC.gold; ctx.textAlign='center'; ctx.font=`bold ${Math.round(PNL*0.50)}px Cairo`;
            ctx.fillText(String(idx), pCX, y+h*0.21);
            ctx.strokeStyle='rgba(199,176,140,0.52)'; ctx.lineWidth=4;
            ctx.beginPath(); ctx.moveTo(x+w-PNL+55,y+h*0.27); ctx.lineTo(x+w-55,y+h*0.27); ctx.stroke();
            ctx.save(); ctx.strokeStyle='rgba(199,176,140,0.80)'; ctx.lineWidth=Math.round(h*0.006); ctx.lineCap='round';
            ctx.beginPath(); ctx.arc(pCX,y+h*0.43,h*0.042,0,Math.PI*2); ctx.stroke();
            ctx.beginPath(); ctx.arc(pCX,y+h*0.49,h*0.078,Math.PI,Math.PI*2,true); ctx.stroke(); ctx.restore();
            ctx.fillStyle='rgba(249,249,249,0.55)'; ctx.font=`normal ${Math.round(h*0.044)}px Cairo`;
            ctx.fillText('منسق التدريب', pCX, y+h*0.64);
            ctx.strokeStyle='rgba(199,176,140,0.44)'; ctx.lineWidth=2.5;
            ctx.beginPath(); ctx.moveTo(x+w-PNL+55,y+h*0.67); ctx.lineTo(x+w-55,y+h*0.67); ctx.stroke();
            const cnFs20=Math.round(h*0.078);
            ctx.fillStyle='#FFFFFF'; ctx.font=`bold ${cnFs20}px Cairo`;
            wrapTextSimple(ctx,c.sp||'—',PNL-50).slice(0,2).forEach((l,li)=>ctx.fillText(l,pCX,y+h*0.76+li*cnFs20*1.35));

            // ── المحتوى (iOS method: احسب أولاً ثم مركز) ─────────
            const bW = w - PNL - PAD*1.5;
            const cX = x + bW/2;
            const ttlFs = Math.round(h * 0.112);
            const dtFs  = Math.round(h * 0.050);
            const chH   = Math.round(h * 0.320);
            const dtGap = Math.round(h * 0.030);
            const chGap = Math.round(h * 0.035);

            const tRes = wrapTextSmart(ctx, c.n, bW-PAD*3, 2, ttlFs);
            const ttlH = tRes.lines.length * tRes.lineHeight;
            const total = ttlH + dtGap + dtFs*1.6 + chGap + chH;
            let cy = y + (h - total) / 2;

            ctx.fillStyle=TC.ink; ctx.textAlign='center'; ctx.font=`bold ${tRes.fontSize}px Cairo`;
            tRes.lines.forEach((l,i)=>ctx.fillText(l, cX, cy+tRes.fontSize+i*tRes.lineHeight));
            cy += ttlH + dtGap;
            nDate(ctx, c, cX, cy+dtFs*0.7, bW-PAD*2, dtFs);
            cy += Math.round(dtFs*1.6) + chGap;
            nChips(ctx, c, x+PAD/2, cy, bW-PAD, chH);
        }

        // ── Template 21: العمليات الذكية — لوحة يسار ─────────────
        function drawRef21ScreenCard(ctx, list, stats, pn, tp, si, logo) {
            tBg(ctx);
            let y = tHdr(ctx, logo, pn, tp) + 55;
            y = tSts(ctx, stats, y) + 65;
            y = tIntr(ctx, y) + 30;
            const n=list.length||1, gap=48;
            const cH=tCardH(y, n, gap);
            for(let i=0;i<n;i++){if(list[i])t21(ctx,list[i],T_MX,y,T_W-T_MX*2,cH,si+i+1);y+=cH+gap;}
            tFtr(ctx);
        }

        function t21(ctx, c, x, y, w, h, idx) {
            const PNL = Math.round(w * 0.192);  // لوحة المنسق يسار
            const PAD = Math.round(h * 0.032);

            // ── البطاقة ───────────────────────────────────────────
            tSh(ctx,46,14); ctx.fillStyle=TC.bgCard; drawRoundedRect(ctx,x,y,w,h,28); tCl(ctx);
            ctx.fillStyle=TC.gold; ctx.fillRect(x,y,w,8);
            ctx.fillStyle=TC.gold; ctx.fillRect(x+w-10,y+22,10,h-44);
            ctx.strokeStyle=TC.line; ctx.lineWidth=3; ctx.strokeRect(x,y,w,h);

            // رقم watermark يمين (خفيف — "01" أسلوب)
            ctx.save(); ctx.globalAlpha=0.038; ctx.fillStyle=TC.gold;
            ctx.textAlign='right'; ctx.font=`bold ${Math.round(h*0.42)}px Cairo`;
            ctx.fillText(String(idx).padStart(2,'0'), x+w-PAD*2, y+h-PAD*2); ctx.restore();

            // ── لوحة المنسق يسار ──────────────────────────────────
            const coX=x+PAD, coY=y+PAD, coH=h-PAD*2;
            tSh(ctx,18,6,'rgba(26,68,69,0.11)');
            ctx.fillStyle=TC.green; drawRoundedRect(ctx,coX,coY,PNL,coH,22); tCl(ctx);
            ctx.fillStyle=TC.gold; ctx.fillRect(coX+PNL-8,coY+14,8,coH-28);
            const piX=coX+PNL/2;
            ctx.save(); ctx.strokeStyle='rgba(199,176,140,0.80)'; ctx.lineWidth=Math.round(h*0.006); ctx.lineCap='round';
            ctx.beginPath(); ctx.arc(piX,coY+coH*0.30,coH*0.080,0,Math.PI*2); ctx.stroke();
            ctx.beginPath(); ctx.arc(piX,coY+coH*0.30+coH*0.12,coH*0.12,Math.PI,Math.PI*2,true); ctx.stroke(); ctx.restore();
            ctx.fillStyle='rgba(249,249,249,0.55)'; ctx.textAlign='center'; ctx.font=`normal ${Math.round(h*0.044)}px Cairo`;
            ctx.fillText('اسم منسق التدريب', piX, coY+coH*0.57);
            ctx.strokeStyle='rgba(199,176,140,0.50)'; ctx.lineWidth=3;
            ctx.beginPath(); ctx.moveTo(coX+35,coY+coH*0.62); ctx.lineTo(coX+PNL-35,coY+coH*0.62); ctx.stroke();
            const cnFs21=Math.round(h*0.076);
            ctx.fillStyle='#FFFFFF'; ctx.font=`bold ${cnFs21}px Cairo`;
            wrapTextSimple(ctx,c.sp||'—',PNL-44).slice(0,2).forEach((l,li)=>ctx.fillText(l,piX,coY+coH*0.73+li*cnFs21*1.35));

            // ── المحتوى (iOS method) ──────────────────────────────
            const cntX = coX + PNL + PAD;
            const cntW = w - PNL - PAD*4 - 10;
            const cX   = cntX + cntW/2;
            const ttlFs = Math.round(h * 0.112);
            const dtFs  = Math.round(h * 0.050);
            const chH   = Math.round(h * 0.320);
            const dtGap = Math.round(h * 0.030);
            const chGap = Math.round(h * 0.035);

            const tRes = wrapTextSmart(ctx, c.n, cntW-PAD*2, 2, ttlFs);
            const ttlH = tRes.lines.length * tRes.lineHeight;
            const total = ttlH + dtGap + dtFs*1.6 + chGap + chH;
            let cy = y + (h - total) / 2;

            ctx.fillStyle=TC.ink; ctx.textAlign='center'; ctx.font=`bold ${tRes.fontSize}px Cairo`;
            tRes.lines.forEach((l,i)=>ctx.fillText(l, cX, cy+tRes.fontSize+i*tRes.lineHeight));
            cy += ttlH + dtGap;
            nDate(ctx, c, cX, cy+dtFs*0.7, cntW-PAD*2, dtFs);
            cy += Math.round(dtFs*1.6) + chGap;
            nChips(ctx, c, cntX, cy, cntW, chH);
        }

        // ==========================================================
        // iOS HD — قالب iOS بجودة فائقة (6480 × 11520)
        // نفس تناسب iOS الأصلي × 3 مع إيقونات خطية ذهبية
        // ==========================================================
        function drawIosHDScreenCard(ctx, list, stats, pn, tp, si, logo) {
            const S = 3;
            const W = 2160*S, H = 3840*S; // 6480 × 11520

            // ── Background ─────────────────────────────────────────
            ctx.fillStyle = '#f5f5f7'; ctx.fillRect(0,0,W,H);

            // ── Header ─────────────────────────────────────────────
            const hPad = 60*S;
            const lTW = 1000*S;
            let lH = 0;
            if(logo&&logo.width>0) lH = lTW*(logo.height/logo.width);
            const tFS = 100*S;
            const hH = hPad + (lH>0?lH+40*S:0) + tFS + hPad;
            ctx.fillStyle = UNI.green; ctx.fillRect(0,0,W,hH);
            ctx.fillStyle = UNI.gold;  ctx.fillRect(0,hH-15*S,W,15*S);
            let hy = hPad;
            if(logo&&logo.width>0){ctx.drawImage(logo,(W-lTW)/2,hy,lTW,lH);hy+=lH+40*S;}
            ctx.fillStyle=UNI.gold; ctx.textAlign='center'; ctx.font=`bold ${tFS}px Cairo`;
            ctx.fillText('جدول الدورات التدريبية',W/2,hy+tFS);
            ctx.fillStyle='rgba(255,255,255,0.55)'; ctx.font=`normal ${36*S}px Cairo`;
            ctx.fillText(`صفحة ${pn} من ${tp}`,W/2,hy+tFS+52*S);

            // ── Stats ───────────────────────────────────────────────
            const sY=hH+48*S, sBH=200*S, sG=34*S, sMX=78*S;
            const sBW=(W-sMX*2-sG*3)/4;
            const sIt=[['إجمالي',stats.t],['جديدة',stats.n],['مستمرة',stats.c],['خارجية',stats.i]];
            sIt.forEach((it,i)=>{
                const bx=sMX+i*(sBW+sG);
                ctx.shadowColor='rgba(0,0,0,0.10)'; ctx.shadowBlur=20*S; ctx.shadowOffsetY=5*S;
                ctx.fillStyle='#ffffff'; drawRoundedRect(ctx,bx,sY,sBW,sBH,20*S);
                ctx.shadowColor='transparent'; ctx.shadowBlur=0; ctx.shadowOffsetY=0;
                ctx.strokeStyle=i===0?UNI.gold:'#e5e5ea'; ctx.lineWidth=(i===0?4:2)*S;
                ctx.strokeRect(bx,sY,sBW,sBH);
                ctx.fillStyle=i===0?UNI.green:'#1d1d1f'; ctx.textAlign='center'; ctx.textBaseline='middle';
                ctx.font=`bold ${90*S}px Cairo`; ctx.fillText(String(it[1]),bx+sBW/2,sY+sBH*0.47);
                ctx.fillStyle='#8e8e93'; ctx.font=`normal ${40*S}px Cairo`; ctx.fillText(it[0],bx+sBW/2,sY+sBH*0.78);
                ctx.textBaseline='alphabetic';
            });

            // ── Intro ───────────────────────────────────────────────
            const iY=sY+sBH+38*S;
            let nY=iY;
            const iText=getScreenIntroText();
            if(iText){
                ctx.fillStyle='#475569'; ctx.textAlign='center'; ctx.font=`normal ${44*S}px Cairo`;
                wrapTextSimple(ctx,iText,W-200*S).slice(0,2).forEach(l=>{ctx.fillText(l,W/2,nY);nY+=58*S;});
            }
            nY+=28*S;

            // ── Cards ───────────────────────────────────────────────
            const cGap=28*S, fH=128*S;
            const cH=Math.floor((H-nY-fH-cGap*(list.length-1))/list.length);
            for(let i=0;i<list.length;i++){
                if(list[i]) iosHDCard(ctx,list[i],78*S,nY,W-156*S,cH,si+i+1,S);
                nY+=cH+cGap;
            }

            // ── Footer ──────────────────────────────────────────────
            ctx.fillStyle='#f5f5f7'; ctx.fillRect(0,H-fH,W,fH);
            ctx.strokeStyle='#e0e0e0'; ctx.lineWidth=2*S;
            ctx.beginPath(); ctx.moveTo(0,H-fH); ctx.lineTo(W,H-fH); ctx.stroke();
            ctx.fillStyle='#8e8e93'; ctx.textAlign='center';
            ctx.font=`normal ${34*S}px Cairo`;
            ctx.fillText('وكالة التدريب بجامعة نايف العربية للعلوم الأمنية',W/2,H-fH/2+14*S);
        }

        function iosHDCard(ctx, c, x, y, w, h, idx, S) {
            const pad = 50*S, rad = 40*S;
            // shadow + card
            ctx.shadowColor='rgba(0,0,0,0.13)'; ctx.shadowBlur=38*S; ctx.shadowOffsetY=10*S;
            ctx.fillStyle='#ffffff'; drawRoundedRect(ctx,x,y,w,h,rad);
            ctx.shadowColor='transparent'; ctx.shadowBlur=0; ctx.shadowOffsetY=0;
            const cTxt='#1d1d1f', cX=x+w/2;

            // sizes (all relative to S)
            const ttlFS=70*S, dtFS=40*S, rBH=96*S, rGp=24*S;
            const spFS=42*S, ttlDG=34*S, dtIG=32*S, iGsS=30*S;

            const res=wrapTextSmart(ctx,c.n,w-pad*2,2,ttlFS);
            const ttlH=res.lines.length*res.lineHeight;
            const spPH=spFS+34*S;
            const total=ttlH+ttlDG+dtFS+dtIG+(rBH*2+rGp)+iGsS+spPH;
            let cY=y+(h-total)/2;

            // title
            ctx.fillStyle=cTxt; ctx.textAlign='center'; ctx.font=`bold ${res.fontSize}px Cairo`;
            res.lines.forEach((l,i)=>ctx.fillText(l,cX,cY+res.fontSize+i*res.lineHeight));
            cY+=ttlH+ttlDG;

            // date
            ctx.fillStyle='#64748b'; ctx.font=`normal ${dtFS}px Cairo`;
            ctx.fillText(`${c.s||'—'} - ${c.e||'—'}`,cX,cY+dtFS);
            cY+=dtFS+dtIG;

            // info boxes (2 rows: 3+2, RTL order)
            const cW=w-pad*2;
            const tBG=22*S, bBG=22*S;
            const tBW=(cW-tBG*2)/3, bBW=(cW-bBG)/2;
            const lFS=22*S, vFS=32*S, icS=36*S, icG=12*S;
            const icCol='#C7B08C', lbCol='#8e8e93';

            const ioBox=(bx,by,bw,bh,ic,lbl,val)=>{
                ctx.fillStyle='#f5f5f7'; drawRoundedRect(ctx,bx,by,bw,bh,18*S);
                const hY=by+bh*0.32;
                ctx.textBaseline='middle'; ctx.font=`normal ${lFS}px Cairo`;
                const lbW=ctx.measureText(lbl).width;
                const gW=lbW+icG+icS;
                const gRX=bx+(bw+gW)/2;
                const icCX=gRX-icS/2, lbCX=gRX-icS-icG-lbW/2;
                iosDrawIcon(ctx,ic,icCX,hY,icS,icCol);
                ctx.fillStyle=lbCol; ctx.textAlign='center'; ctx.fillText(lbl,lbCX,hY);
                ctx.fillStyle=cTxt; ctx.font=`bold ${vFS}px Cairo`;
                let sv=String(val||'—'); while(ctx.measureText(sv).width>bw-36*S&&sv.length>0)sv=sv.slice(0,-1);
                if(sv!==String(val||'—'))sv+='…';
                ctx.fillText(sv,bx+bw/2,by+bh*0.74);
                ctx.textBaseline='alphabetic';
            };

            const r1=[{ic:'location',lb:'مكان التنفيذ',v:c.loc},{ic:'status',lb:'الحالة',v:c.st},{ic:'period',lb:'الفترة',v:c.p}];
            const r2=[{ic:'room',lb:'القاعة',v:c.r},{ic:'floor',lb:'الدور',v:shouldHideFloorValue(c.f)?'—':c.f}];
            r1.forEach((it,i)=>{const bx=x+pad+cW-tBW-i*(tBW+tBG);ioBox(bx,cY,tBW,rBH,it.ic,it.lb,it.v);});
            cY+=rBH+rGp;
            r2.forEach((it,i)=>{const bx=x+pad+cW-bBW-i*(bBW+bBG);ioBox(bx,cY,bBW,rBH,it.ic,it.lb,it.v);});
            cY+=rBH+iGsS;

            // coordinator
            const spTxt=`اسم منسق التدريب: ${c.sp||'—'}`;
            ctx.font=`bold ${spFS}px Cairo`;
            const spW=Math.min(w-pad*2,ctx.measureText(spTxt).width+90*S);
            ctx.fillStyle='rgba(128,47,45,0.07)';
            drawRoundedRect(ctx,cX-spW/2,cY+6*S,spW,spFS+34*S,18*S);
            ctx.fillStyle=UNI.supRed; ctx.textAlign='center'; ctx.fillText(spTxt,cX,cY+spFS);

            // index badge (top-left of card in canvas = leading in RTL)
            const bdR=28*S, bdX=x+pad*0.7, bdY=y+pad*0.7;
            ctx.strokeStyle=UNI.green; ctx.lineWidth=4*S;
            ctx.beginPath(); ctx.arc(bdX,bdY,bdR,0,Math.PI*2); ctx.stroke();
            ctx.fillStyle=UNI.green; ctx.textAlign='center'; ctx.font=`bold ${22*S}px Cairo`;
            ctx.fillText(String(idx).padStart(2,'0'),bdX,bdY+9*S);
        }

        function iosDrawIcon(ctx, type, cx, cy, s, col) {
            ctx.save(); ctx.strokeStyle=col; ctx.lineWidth=Math.max(3,s*0.12);
            ctx.lineCap='round'; ctx.lineJoin='round';
            if(type==='location'){
                ctx.beginPath(); ctx.arc(cx,cy-s*0.12,s*0.18,0,Math.PI*2); ctx.stroke();
                ctx.beginPath(); ctx.moveTo(cx,cy+s*0.42);
                ctx.lineTo(cx-s*0.18,cy+s*0.08);
                ctx.quadraticCurveTo(cx-s*0.28,cy-s*0.1,cx,cy-s*0.3);
                ctx.quadraticCurveTo(cx+s*0.28,cy-s*0.1,cx+s*0.18,cy+s*0.08);
                ctx.closePath(); ctx.stroke();
            } else if(type==='status'){
                ctx.beginPath(); ctx.arc(cx,cy,s*0.3,0,Math.PI*2); ctx.stroke();
                ctx.beginPath(); ctx.moveTo(cx-s*0.12,cy+s*0.02);
                ctx.lineTo(cx-s*0.02,cy+s*0.14); ctx.lineTo(cx+s*0.16,cy-s*0.10); ctx.stroke();
            } else if(type==='period'){
                ctx.beginPath(); ctx.arc(cx,cy,s*0.3,0,Math.PI*2); ctx.stroke();
                ctx.beginPath(); ctx.moveTo(cx,cy); ctx.lineTo(cx,cy-s*0.14);
                ctx.moveTo(cx,cy); ctx.lineTo(cx+s*0.12,cy+s*0.08); ctx.stroke();
            } else if(type==='room'){
                ctx.strokeRect(cx-s*0.28,cy-s*0.22,s*0.56,s*0.44);
                ctx.beginPath(); ctx.moveTo(cx-s*0.34,cy-s*0.22);
                ctx.lineTo(cx,cy-s*0.42); ctx.lineTo(cx+s*0.34,cy-s*0.22); ctx.stroke();
                ctx.beginPath(); ctx.moveTo(cx,cy+s*0.22); ctx.lineTo(cx,cy-s*0.02); ctx.stroke();
            } else {
                ctx.beginPath();
                ctx.moveTo(cx-s*0.28,cy+s*0.18); ctx.lineTo(cx-s*0.08,cy+s*0.18);
                ctx.lineTo(cx-s*0.08,cy); ctx.lineTo(cx+s*0.12,cy);
                ctx.lineTo(cx+s*0.12,cy-s*0.18); ctx.lineTo(cx+s*0.3,cy-s*0.18); ctx.stroke();
            }
            ctx.restore();
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
            const baseTitleSize = isVibrantScreen ? 74 : (isScreen ? 70 : 42);
            const dateFontSize = isVibrantScreen ? 42 : (isScreen ? 40 : 25);
            const gridH = isIosTemplate ? (isScreen ? 238 : 144) : (isVibrantScreen ? 186 : (isScreen ? 174 : 108));
            const supFontSize = isIosTemplate ? (isScreen ? 42 : 27) : (isVibrantScreen ? 47 : (isScreen ? 40 : 25));
            const gapSize = isVibrantScreen ? 48 : (isScreen ? 34 : 18);
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
                
                if (currentTemplate === 4) { ctx.font = `normal 32px Cairo`; ctx.fillText(label, gridStartX + boxW/2, currentY + gridH * 0.30); }
                else { ctx.font = `normal ${isVibrantScreen ? 28*sc : (isScreen ? 26 : 17) * sc}px Cairo`; ctx.textAlign = 'center'; ctx.textBaseline = 'middle'; ctx.fillText(label, gridStartX + boxW/2, currentY + gridH * 0.28); }
                
                ctx.fillStyle = cardText; ctx.font = `normal ${isVibrantScreen ? 38*sc : (isScreen ? 34 : 22) * sc}px Cairo`;
                let safeValue = String(displayValue);
                while (ctx.measureText(safeValue).width > boxW - 24*sc && safeValue.length > 0) safeValue = safeValue.substring(0, safeValue.length - 1);
                if (safeValue !== String(displayValue)) safeValue += '…';
                ctx.fillText(safeValue, gridStartX + boxW/2, currentY + gridH * 0.72);
                gridStartX += boxW + boxGap;
            };

            const drawIosInfoBoxAt = (boxX, boxY, boxWLocal, boxHLocal, iconType, label, value) => {
                const displayValue = value || '-';
                ctx.fillStyle = '#f5f5f7';
                drawRoundedRect(ctx, boxX, boxY, boxWLocal, boxHLocal, 18*sc);

                const labelFontSize = isScreen ? 22 : 13;
                const valueFontSize = isScreen ? 32 : 20;
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
                ctx.font = `normal ${valueFontSize}px Cairo`;
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
            const supervisorText = `اسم منسق التدريب: ${c.sp}`;
            ctx.font = `normal ${supFontSize}px Cairo`;
            const supervisorW = Math.min(w - padding * 2, ctx.measureText(supervisorText).width + 90 * sc);
            ctx.fillStyle = currentTemplate === 6 ? 'rgba(255,255,255,0.82)' : (currentTemplate === 5 ? 'rgba(128,47,45,0.08)' : 'rgba(128,47,45,0.07)');
            drawRoundedRect(ctx, centerX - supervisorW / 2, currentY + 6 * sc, supervisorW, supFontSize + 34 * sc, 18 * sc);

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