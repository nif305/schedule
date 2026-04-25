// --- 1. Setup & Helpers ---
        let rows = [];
        let currentTemplate = 1;
        let logoImage = null; 
        let lastBlob = null; 
        let lastZipBlob = null; 

        const FIXED_LOGO_URL = "./public/assets/logo-footer.png";
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

        const KEYS = { n: "اسم الدورة", s: "بداية", e: "نهاية", r: "قاعة", f: "طابق", p: "فترة", st: "حالة", sp: "مشرف", t: "نوع" };
        const ARCH_KEY = "nfdp_archive_v42";
        const MEMORY_KEY = "nfdp_memory_v1";
        let lastUsed = {};

        function showToast(message, type = 'info') { const container = document.getElementById('toast-container'); const toast = document.createElement('div'); toast.className = `toast ${type}`; toast.innerText = message; container.appendChild(toast); setTimeout(() => { toast.style.animation = 'fadeOut 0.3s ease-out forwards'; setTimeout(() => toast.remove(), 300); }, 3000); }

        function parseExcelDate(value) {
            if (!value) return '';
            const formatDate = (dateObj) => { let year = dateObj.getFullYear(); let month = (dateObj.getMonth() + 1).toString().padStart(2, '0'); let day = dateObj.getDate().toString().padStart(2, '0'); return `${year}-${month}-${day}`; };
            if (value instanceof Date) return formatDate(value);
            if (typeof value === 'number') { const date = new Date((value - 25569) * 86400 * 1000); return formatDate(date); }
            if (typeof value === 'string') { if (/^\d{4}-\d{2}-\d{2}$/.test(value)) return value; const date = new Date(value); if (!isNaN(date.getTime())) return formatDate(date); }
            return '';
        }
        
        // Helper to format date nicely (YYYY/MM/DD)
        function formatDateShort(dateStr) {
            if(!dateStr) return '';
            // Replace - with /
            return dateStr.replace(/-/g, '/');
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

        document.addEventListener('DOMContentLoaded', () => { addRow(); initWeekNumber(); loadDefaultLogo(); loadMemory(); checkShareSupport(); });

        // --- 2. UI Logic ---
        function initWeekNumber() { try { const archive = JSON.parse(localStorage.getItem(ARCH_KEY) || '{}'); const weeks = Object.keys(archive).map(Number).filter(n => !isNaN(n) && n > 0); let maxWeek = 0; if (weeks.length > 0) maxWeek = Math.max(...weeks); document.getElementById('w-id').value = maxWeek + 1; } catch(e) {} }
        function showPage(pageId) { document.querySelectorAll('.page-container').forEach(el => el.classList.remove('active')); document.getElementById('page-' + pageId).classList.add('active'); if(pageId === 'archive') refreshArchiveView(); }
        function resetAndShowGenerator() { if(rows.length > 1 && !confirm("هل تريد بدء جدول جديد؟")) return; clearCurrentData(); showPage('generator'); document.querySelectorAll('.nav-tab').forEach(el => el.classList.remove('active')); document.querySelector('.nav-tab').classList.add('active'); }
        function clearCurrentData() { rows = []; document.getElementById('rows-box').innerHTML = ''; addRow(); updateIdx(); showToast("تم مسح البيانات", 'info'); }
        function selectTemplate(id) { currentTemplate = id; document.querySelectorAll('.template-thumb').forEach(el => el.classList.remove('selected')); document.getElementById('tpl-' + id).classList.add('selected'); }
        function toggleLoader(show, text = "جاري المعالجة...") { const el = document.getElementById('loader'); document.getElementById('loader-text').innerText = text; if(show) { el.classList.remove('hidden'); el.classList.add('flex'); } else { el.classList.add('hidden'); el.classList.remove('flex'); } }

        function loadMemory() { const m = localStorage.getItem(MEMORY_KEY); if(m) lastUsed = JSON.parse(m); }
        function saveMemory(key, val) { lastUsed[key] = val; localStorage.setItem(MEMORY_KEY, JSON.stringify(lastUsed)); }

        // الشعار الرسمي ثابت ويحمّل تلقائيًا من public/assets/logo-footer.png

        function addRow(d = {}) { 
            const id = Date.now() + Math.random(); 
            const o = { id: id, n: d.n || "", s: d.s || "", e: d.e || "", r: d.r || lastUsed.r || DDL.room[0], f: d.f || lastUsed.f || DDL.floor[0], p: d.p || lastUsed.p || DDL.period[0], st: d.st || lastUsed.st || DDL.status[0], sp: d.sp || lastUsed.sp || DDL.sup[0], t: d.t || lastUsed.t || DDL.type[0] }; 
            rows.push(o); renderRow(o); 
        }
        
        function renderRow(o) {
            const el = document.createElement('div');
            el.id = `r-${o.id}`;
            el.className = "bg-white p-2 rounded-lg border hover:border-[#2c6060] transition-colors relative group";
            const opts = (arr, sel) => arr.map(x => `<option value="${x}" ${x===sel?'selected':''}>${x}</option>`).join('');
            el.innerHTML = `
                <div class="grid grid-cols-12 gap-2 items-center">
                    <div class="col-span-1 text-center text-slate-300 font-bold text-xs idx">${rows.length}</div>
                    <div class="col-span-11 md:col-span-4"><input type="text" placeholder="اسم الدورة" value="${o.n}" onchange="upd(${o.id},'n',this.value)" class="form-input border-transparent focus:border-[#2c6060]"></div>
                    <div class="col-span-6 md:col-span-2"><select onchange="upd(${o.id},'t',this.value)" class="form-input text-xs">${opts(DDL.type, o.t)}</select></div>
                    <div class="col-span-6 md:col-span-2"><select onchange="upd(${o.id},'st',this.value)" class="form-input text-xs">${opts(DDL.status, o.st)}</select></div>
                    <div class="col-span-11 md:col-span-2"><select onchange="upd(${o.id},'sp',this.value)" class="form-input text-xs">${opts(DDL.sup, o.sp)}</select></div>
                    <div class="col-span-1 flex flex-col gap-1 items-center"><div class="w-6 h-6 flex items-center justify-center cursor-pointer bg-red-50 text-red-500 rounded hover:bg-red-100 text-sm font-bold" onclick="delRow(${o.id})">×</div></div>
                </div>
                <div class="grid grid-cols-12 gap-2 items-center mt-1">
                    <div class="hidden md:block col-span-1"></div>
                    <div class="col-span-6 md:col-span-2"><input type="date" value="${o.s}" onchange="upd(${o.id},'s',this.value)" class="form-input text-xs"></div>
                    <div class="col-span-6 md:col-span-2"><input type="date" value="${o.e}" onchange="upd(${o.id},'e',this.value)" class="form-input text-xs"></div>
                    <div class="col-span-6 md:col-span-2"><select onchange="upd(${o.id},'r',this.value)" class="form-input text-xs">${opts(DDL.room, o.r)}</select></div>
                    <div class="col-span-6 md:col-span-2"><select onchange="upd(${o.id},'f',this.value)" class="form-input text-xs">${opts(DDL.floor, o.f)}</select></div>
                    <div class="col-span-12 md:col-span-2"><select onchange="upd(${o.id},'p',this.value)" class="form-input text-xs">${opts(DDL.period, o.p)}</select></div>
                </div>`;
            document.getElementById('rows-box').appendChild(el);
        }
        
        function upd(id, k, v) { let o = rows.find(x => x.id === id); if(o) o[k] = v; saveMemory(k, v); }
        function delRow(id) { rows = rows.filter(x => x.id !== id); document.getElementById(`r-${id}`).remove(); updateIdx(); if (rows.length === 0) addRow(); }
        function updateIdx() { document.querySelectorAll('.idx').forEach((n, i) => n.innerText = i+1); }

        // --- 3. Excel ---
        function dlTpl() { const sample = ["دورة تجريبية", "2023-10-01", "2023-10-05", "LAB 1", "الأرضي", "صباحي", "جديدة", "نايف الشهراني", "داخلية"]; const ws = XLSX.utils.aoa_to_sheet([[KEYS.n, KEYS.s, KEYS.e, KEYS.r, KEYS.f, KEYS.p, KEYS.st, KEYS.sp, KEYS.t], sample]); const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "Template"); XLSX.writeFile(wb, "NFDP_Template.xlsx"); showToast("تم تحميل القالب", 'success'); }
        function rdExcel(e) {
            const f = e.target.files[0]; if(!f) return;
            const rd = new FileReader();
            rd.onload = function(ev) {
                try { const data = new Uint8Array(ev.target.result); const workbook = XLSX.read(data, {type:'array', cellDates: true}); const sheet = workbook.Sheets[workbook.SheetNames[0]]; const json = XLSX.utils.sheet_to_json(sheet); document.getElementById('rows-box').innerHTML = ''; rows = []; json.forEach(row => { addRow({ n: row[KEYS.n] || '', s: parseExcelDate(row[KEYS.s]), e: parseExcelDate(row[KEYS.e]), r: row[KEYS.r] || DDL.room[0], f: row[KEYS.f] || DDL.floor[0], p: row[KEYS.p] || DDL.period[0], st: row[KEYS.st] || DDL.status[0], sp: row[KEYS.sp] || DDL.sup[0], t: row[KEYS.t] || DDL.type[0] }); }); updateIdx(); showToast(`تم استيراد ${rows.length} دورات`, 'success'); e.target.value = ''; } catch(err) { showToast("خطأ في قراءة الملف: " + err.message, 'error'); }
            };
            rd.readAsArrayBuffer(f);
        }

        // --- 4. Archive ---
        function autoSave() { try { const week = document.getElementById('w-id').value; const owner = document.getElementById('owner-select').value; const archive = JSON.parse(localStorage.getItem(ARCH_KEY) || '{}'); archive[week] = { rows: rows, owner: owner }; localStorage.setItem(ARCH_KEY, JSON.stringify(archive)); } catch(e) {} }
        function loadFromArchive(week) { const archive = JSON.parse(localStorage.getItem(ARCH_KEY) || '{}'); if(archive[week]) { const data = archive[week]; rows = []; document.getElementById('rows-box').innerHTML = ''; (data.rows || data).forEach(r => addRow(r)); document.getElementById('w-id').value = week; document.getElementById('owner-select').value = data.owner || "نايف الشهراني"; updateIdx(); showPage('generator'); showToast(`تم تحميل الأسبوع ${week}`, 'info'); } }
        function deleteFromArchive(week) { if(confirm(`حذف أرشيف الأسبوع ${week}؟`)) { const archive = JSON.parse(localStorage.getItem(ARCH_KEY) || '{}'); delete archive[week]; localStorage.setItem(ARCH_KEY, JSON.stringify(archive)); refreshArchiveView(); initWeekNumber(); showToast("تم الحذف", 'info'); } }
        function refreshArchiveView() { const container = document.getElementById('archive-list'); const noData = document.getElementById('no-archive'); const archive = JSON.parse(localStorage.getItem(ARCH_KEY) || '{}'); const keys = Object.keys(archive); container.innerHTML = ''; if(keys.length === 0) { noData.classList.remove('hidden'); return; } noData.classList.add('hidden'); keys.sort((a,b) => b - a).forEach(key => { const entry = archive[key]; const owner = entry.owner || 'غير محدد'; const count = (entry.rows || entry).length; const div = document.createElement('div'); div.className = "bg-slate-50 p-4 rounded-lg border hover:shadow-md transition-shadow"; div.innerHTML = `<div class="flex justify-between items-center mb-2"><h4 class="font-bold text-[#2c6060] text-lg">الأسبوع ${key}</h4><span class="text-xs bg-slate-200 px-2 py-1 rounded">${count} دورات</span></div><div class="text-xs text-slate-500 mb-3">بواسطة: <span class="font-bold text-slate-700">${owner}</span></div><div class="flex gap-2 mt-4"><button onclick="loadFromArchive('${key}')" class="flex-1 btn-main text-xs">📝 تعديل</button><button onclick="deleteFromArchive('${key}')" class="btn-outline text-xs text-red-500 border-red-500 hover:bg-red-50">🗑️ حذف</button></div>`; container.appendChild(div); }); }

        // --- 5. Save Logic ---
        
        async function startScreenSave() {
            const valid = rows.filter(x => x.n.trim() !== ""); if(valid.length === 0) { showToast("لا توجد بيانات!", 'error'); return; }
            toggleLoader(true, "جاري التوليد...");
            try {
                if(document.fonts && document.fonts.ready) await document.fonts.ready;
                autoSave(); const week = document.getElementById('w-id').value; const stats = getStats(valid); const zip = new JSZip();
                const chunksS = chunk(valid, 4); 
                for(let i=0; i<chunksS.length; i++) { 
                    const canvas = document.createElement('canvas'); canvas.width = 2160; canvas.height = 3840; 
                    const ctx = canvas.getContext('2d'); drawScreenCard(ctx, chunksS[i], stats, i+1, chunksS.length, i*4, logoImage); 
                    const blob = await new Promise(resolve => canvas.toBlob(resolve, 'image/jpeg', 0.95)); zip.file(`Week-${week}-Screen-${i+1}.jpg`, blob); 
                }
                const content = await zip.generateAsync({type: "blob"});
                lastZipBlob = content; 
                saveAs(content, `NFDP_Week_${week}_Screens.zip`); showToast("تم الحفظ!", 'success');
            } catch (err) { console.error(err); showToast("خطأ: " + err.message, 'error'); } finally { toggleLoader(false); }
        }

        async function startAgentReport() {
            const valid = rows.filter(x => x.n.trim() !== ""); if(valid.length === 0) { showToast("لا توجد بيانات!", 'error'); return; }
            toggleLoader(true, "جاري التوليد...");
            try {
                if(document.fonts && document.fonts.ready) await document.fonts.ready;
                autoSave(); const week = document.getElementById('w-id').value; const stats = getStats(valid);
                const canvas = document.createElement('canvas'); canvas.width = 1080; canvas.height = 1920; const ctx = canvas.getContext('2d');
                drawAgentReport(ctx, valid, stats, logoImage);
                const blob = await new Promise(resolve => canvas.toBlob(resolve, 'image/jpeg', 0.9));
                lastBlob = blob; 
                saveAs(blob, `NFDP_Week_${week}_Agent_Report.jpg`);
                showToast("تم التوليد!", 'success');
            } catch (err) { console.error(err); showToast("خطأ: " + err.message, 'error'); } finally { toggleLoader(false); }
        }

        // --- Share Logic ---
        function checkShareSupport() { 
            const btnScreen = document.getElementById('btn-share-screen'); 
            const btnReport = document.getElementById('btn-share-report');
            btnScreen.classList.remove('hidden'); 
            btnReport.classList.remove('hidden'); 
        }

        async function shareScreenImages() {
            if(!lastZipBlob) { showToast("يرجى توليد صور العرض أولاً", 'info'); return; }
            const file = new File([lastZipBlob], "training_screens.zip", { type: "application/zip" });
            if (navigator.share && navigator.canShare && navigator.canShare({ files: [file] })) {
                try { await navigator.share({ files: [file], title: 'جدول العرض', text: 'جداول الدورات التدريبية' }); } catch (error) { console.log('Share cancelled', error); }
            } else { showToast("جهازك لا يدعم مشاركة ملفات ZIP مباشرة", 'error'); }
        }

        async function shareAgentReport() {
            if(!lastBlob) { showToast("يرجى توليد تقرير الوكيل أولاً", 'info'); return; }
            const file = new File([lastBlob], "report.jpg", { type: "image/jpeg" });
            if (navigator.share && navigator.canShare && navigator.canShare({ files: [file] })) {
                try { await navigator.share({ files: [file], title: 'تقرير الوكيل', text: 'جدول الدورات التدريبية' }); } catch (error) { console.log('Share cancelled', error); }
            } else { showToast("المتصفح لا يدعم المشاركة", 'error'); }
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
            drawStatBox(bX, 'دولية', stats.i);

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
                const details = `${item.t}  •  ${item.st}  •  ${item.r}`;
                ctx.textAlign = 'right';
                ctx.fillText(details, W - 100, y + 58);

                // Date on Left
                const sDate = formatDateShort(item.s);
                const eDate = formatDateShort(item.e);
                const dateStr = (sDate && eDate) ? `${sDate} - ${eDate}` : (sDate || eDate);
                ctx.fillStyle = UNI.green;
                ctx.textAlign = 'left';
                ctx.fillText(dateStr, 100, y + 58);

                // Supervisor
                ctx.fillStyle = UNI.supRed; ctx.font = 'bold 16px Cairo';
                ctx.textAlign = 'right';
                ctx.fillText(`👤 ${item.sp}`, W - 100, y + 85);
                
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
            const logoTargetW = 1000; 
            let logoH = 0;
            if (logo && logo.width > 0) { logoH = logoTargetW * (logo.height / logo.width); }
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
            if (logo && logo.width > 0) { const logoX = (W - logoTargetW) / 2; ctx.drawImage(logo, logoX, currentDrawY, logoTargetW, logoH); currentDrawY += logoH + 40; }
            
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
            const introText = "ينفذ برنامج الشراكات الدولية بوزارة الداخلية بالتعاون مع جامعة نايف العربية للعلوم الأمنية البرامج التدريبية التالية:";
            const introLines = wrapTextSimple(ctx, introText, W - 200);
            let introY = headerEndY + 400;
            introLines.forEach(line => { ctx.fillText(line, W/2, introY); introY += 60; });

            let y = introY + 80;
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
            ctx.font = `normal ${40*sc}px Cairo`; ctx.fillText('دولية', startX + (boxW + gap)*3 + boxW/2, Y + boxH/2 + 35*sc);
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

            const baseTitleSize = isScreen ? 65 : 40; const dateFontSize = isScreen ? 35 : 24; const gridH = isScreen ? 150 : 100; const supFontSize = isScreen ? 35 : 24; const gapSize = isScreen ? 25 : 15;
            const fullTitle = c.n + (c.t === 'دولية' ? ' 🌍' : '');
            const result = wrapTextSmart(ctx, fullTitle, w - padding*2, 2, baseTitleSize);
            const titleLines = result.lines; const titleFontSize = result.fontSize; const lineHeight = result.lineHeight;
            const titleH = titleLines.length * lineHeight;
            
            const totalContentH = titleH + gapSize + dateFontSize + gapSize + gridH + gapSize + supFontSize;
            let currentY = y + (h - totalContentH) / 2;
            const centerX = x + w / 2;

            ctx.font = `bold ${titleFontSize}px Cairo`; ctx.fillStyle = cardText; ctx.textAlign = 'center';
            let textY = currentY + titleFontSize;
            titleLines.forEach((line, index) => { ctx.fillText(line, centerX, textY); textY += lineHeight; });
            currentY += titleH + gapSize;

            if (currentTemplate === 6) ctx.fillStyle = 'rgba(1, 101, 100, 0.8)';
            else if(currentTemplate === 2) ctx.fillStyle = UNI.greenMid;
            else if(currentTemplate === 5) ctx.fillStyle = '#555';
            else ctx.fillStyle = '#64748b';
            ctx.font = `normal ${dateFontSize}px Cairo`; ctx.fillText(`${c.s} - ${c.e}`, centerX, currentY + dateFontSize); currentY += dateFontSize + gapSize;

            let gridStartX = x + padding; const boxW = (w - padding*2 - (20*sc*3)) / 4;
            const drawInfoBox = (label, value) => {
                if (currentTemplate === 6) { ctx.fillStyle = 'rgba(1, 101, 100, 0.1)'; drawRoundedRect(ctx, gridStartX, currentY, boxW, gridH, 8*sc); ctx.fillStyle = UNI.green; }
                else if (currentTemplate === 4) { ctx.fillStyle = '#f5f5f7'; drawRoundedRect(ctx, gridStartX, currentY, boxW, gridH, 15*sc); ctx.fillStyle = '#8e8e93'; }
                else if(currentTemplate === 5) {
                    ctx.fillStyle = 'rgba(161, 136, 127, 0.1)'; drawRoundedRect(ctx, gridStartX, currentY, boxW, gridH, 0);
                    ctx.fillStyle = '#5d4037';
                }
                else if(currentTemplate === 2) { ctx.fillStyle = '#f0fdfa'; drawRoundedRect(ctx, gridStartX, currentY, boxW, gridH, 8*sc); ctx.strokeStyle = UNI.gold; ctx.lineWidth = 1; ctx.strokeRect(gridStartX, currentY, boxW, gridH); ctx.fillStyle = UNI.green; }
                else { ctx.fillStyle = '#f8fafc'; drawRoundedRect(ctx, gridStartX, currentY, boxW, gridH, 10*sc); ctx.fillStyle = '#64748b'; }
                
                if (currentTemplate === 4) { ctx.font = `bold 30px Cairo`; ctx.fillText(label, gridStartX + boxW/2, currentY + gridH * 0.35); }
                else { ctx.font = `normal ${(isScreen ? 24 : 16) * sc}px Cairo`; ctx.textAlign = 'center'; ctx.textBaseline = 'middle'; ctx.fillText(label, gridStartX + boxW/2, currentY + gridH * 0.35); }
                
                ctx.fillStyle = cardText; ctx.font = `bold ${(isScreen ? 32 : 20) * sc}px Cairo`; ctx.fillText(value, gridStartX + boxW/2, currentY + gridH * 0.7);
                gridStartX += boxW + (20*sc);
            };
            
            if (currentTemplate === 4) { drawInfoBox('🏛️', c.r); drawInfoBox('🧱', c.f); drawInfoBox('⏰', c.p); drawInfoBox('⚡', c.st); }
            else { drawInfoBox('القاعة', c.r); drawInfoBox('الطابق', c.f); drawInfoBox('الفترة', c.p); drawInfoBox('الحالة', c.st); }
            
            currentY += gridH + gapSize;

            ctx.fillStyle = UNI.supRed;
            ctx.font = `normal ${supFontSize}px Cairo`; ctx.fillText(`المشرف: ${c.sp}`, centerX, currentY + supFontSize);

            const idxSize = (isScreen ? 24 : 12) * sc;
            if (currentTemplate === 6) { ctx.fillStyle = 'rgba(255,255,255,0.3)'; ctx.beginPath(); ctx.arc(x + 50*sc, y + 50*sc, 25*sc, 0, 2 * Math.PI); ctx.fill(); ctx.fillStyle = '#ffffff'; }
            else if (currentTemplate === 4) { ctx.fillStyle = '#f5f5f7'; ctx.beginPath(); ctx.arc(x + 50*sc, y + 50*sc, 25*sc, 0, 2 * Math.PI); ctx.fill(); ctx.fillStyle = '#8e8e93'; }
            else if(currentTemplate === 5) {
                 ctx.fillStyle = '#a1887f'; ctx.fillRect(x + 30*sc, y + 20*sc, 50*sc, 50*sc);
                 ctx.fillStyle = '#ffffff';
            }
            else if(currentTemplate === 2) ctx.fillStyle = UNI.green; else ctx.fillStyle = '#2c6060';
            if(currentTemplate !== 4 && currentTemplate !== 5) { ctx.beginPath(); ctx.arc(x + 50*sc, y + 50*sc, 25*sc, 0, 2 * Math.PI); ctx.fill(); ctx.fillStyle = '#fff'; }
            ctx.font = `bold ${idxSize}px Cairo`; ctx.textAlign = 'center'; ctx.textBaseline = 'middle'; ctx.fillText(idx, x + 55*sc, y + 45*sc);
        }

        function getStats(a) { return { t: a.length, n: a.filter(x=>x.st==='جديدة').length, c: a.filter(x=>x.st==='مستمرة').length, i: a.filter(x=>x.t==='دولية').length }; }
        function chunk(a, s) { const r=[]; for(let i=0; i<a.length; i+=s) r.push(a.slice(i,i+s)); return r; }
