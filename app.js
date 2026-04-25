// ═══════════════════════════════════════════
// LAZY LOADER — xlsx dimuat hanya saat dibutuhkan
// ═══════════════════════════════════════════
function loadXlsx(callback) {
  if (window.XLSX) { callback(); return; }
  const script = document.createElement('script');
  script.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
  script.onload = callback;
  script.onerror = () => alert('Gagal memuat library Excel. Periksa koneksi internet kamu.');
  document.head.appendChild(script);
}

// ═══════════════════════════════════════════
// CONSTANTS — single source of truth
// ═══════════════════════════════════════════
const LS_KEY_REV           = 'taobun_master_final_v7';
const LS_KEY_SOBUN         = 'taobun_loyalty_v1';
const LS_KEY_EXCEL_BASELINE = 'taobun_excel_baseline_v1'; // Data permanen dari import Excel
const LS_KEY_BRANCH        = 'taobun_branch_v1'; // pastikan konsisten
const APP_NAME             = 'TAOBUN Management Suite';

// State: apakah data saat ini berasal dari Excel baseline?
let _xlBaselineLoaded = false;

if (typeof Chart !== 'undefined' && typeof ChartDataLabels !== 'undefined') {
  Chart.register(ChartDataLabels);
}
let funnelChart, growthChart;
let currentMonth = "02", currentYear = "2026";
const months = [
    {v: '01', l: 'Januari'}, {v: '02', l: 'Februari'}, {v: '03', l: 'Maret'},
    {v: '04', l: 'April'}, {v: '05', l: 'Mei'}, {v: '06', l: 'Juni'},
    {v: '07', l: 'Juli'}, {v: '08', l: 'Agustus'}, {v: '09', l: 'September'},
    {v: '10', l: 'Oktober'}, {v: '11', l: 'November'}, {v: '12', l: 'Desember'}
];

// ═══════════════════════════════════════════
// CUSTOM CONFIRM DIALOG (callback-based, no async/await)
// ═══════════════════════════════════════════
let _confirmCallback = null;

function showConfirm({ icon = '⚠️', title, msg, okLabel = 'Lanjutkan' } = {}, callback) {
    _confirmCallback = callback || null;
    const el = document.getElementById('confirmOverlay');
    document.getElementById('confirmIcon').textContent  = icon;
    document.getElementById('confirmTitle').textContent = title;
    document.getElementById('confirmMsg').textContent   = msg;
    document.getElementById('confirmOkBtn').textContent = okLabel;
    // Move to body root to escape any stacking context from parent elements
    document.body.appendChild(el);
    el.classList.add('open');
}
function confirmResolve(result) {
    const el = document.getElementById('confirmOverlay');
    el.classList.remove('open');
    if (_confirmCallback) {
        const cb = _confirmCallback;
        _confirmCallback = null;
        if (result) cb();
    }
}

// ═══════════════════════════════════════════
// INPUT VALIDATION
// ═══════════════════════════════════════════
function validateInput(el, opts = {}) {
    const { min = 0, max = Infinity, required = false } = opts;
    const val = cleanNumber(el.value);
    let err = '';
    if (required && (!el.value || el.value === '0')) err = 'Field ini wajib diisi.';
    else if (val < min) err = `Nilai minimal ${min.toLocaleString('id-ID')}.`;
    else if (val > max) err = `Nilai maksimal ${max.toLocaleString('id-ID')}.`;

    const errEl = el.parentElement?.querySelector('.field-error-msg') ||
                  el.closest('.deck-item')?.querySelector('.field-error-msg');
    if (err) {
        el.classList.add('input-error');
        if (errEl) { errEl.textContent = err; errEl.classList.add('show'); }
        return false;
    } else {
        el.classList.remove('input-error');
        if (errEl) errEl.classList.remove('show');
        return true;
    }
}
function clearValidation(el) {
    el.classList.remove('input-error');
    const errEl = el.parentElement?.querySelector('.field-error-msg') ||
                  el.closest('.deck-item')?.querySelector('.field-error-msg');
    if (errEl) errEl.classList.remove('show');
}

// ═══════════════════════════════════════════
// IMPORT LOADING HELPER
// ═══════════════════════════════════════════
function showImportLoading(msg = 'Memproses data...') {
    document.getElementById('importLoadingTxt').textContent = msg;
    document.getElementById('importLoading').classList.add('show');
}
function hideImportLoading() {
    document.getElementById('importLoading').classList.remove('show');
}

function formatNumber(val) {
  let clean = val.toString().replace(/[^0-9]/g, "");
  return clean.replace(/\B(?=(\d{3})+(?!\d))/g, ".");
}
function cleanNumber(val) { return parseFloat(val.toString().replace(/\./g, "")) || 0; }
function formatShortNumber(num) {
  if (num >= 1000000000) return (num / 1000000000).toFixed(1) + 'M';
  if (num >= 1000000) return (num / 1000000).toFixed(1) + 'Jt';
  return num.toLocaleString('id-ID');
}

function clearIfZero(el) { if (el.value === "0") el.value = ""; }
function restoreIfEmpty(el) {
  if (el.value === "") {
    el.value = "0";
    // Only call the relevant calculation
    if (el.classList.contains('lead-field')) { calculateLeads(); }
    else { calculate(); }
  }
}

function showView(id) {
    document.querySelectorAll('.view-container').forEach(v => v.classList.remove('view-active'));
    const target = document.getElementById(id);
    if (target) target.classList.add('view-active');
    // Simpan view aktif ke localStorage agar bisa di-restore setelah refresh
    try { localStorage.setItem('taobun_active_view', id); } catch(e) {}
    if (id === 'home-view') { buildHomeAlerts(); checkOnboarding(); }
    // Re-render active sobun tab when entering sobun-view
    if (id === 'sobun-view') {
        const activeTab = document.querySelector('.sobun-wrapper .page.active');
        if (activeTab) {
            const tabId = activeTab.id.replace('sb-page-', '');
            if (tabId === 'master')    renderMaster();
            else if (tabId === 'reward')    renderRw();
            else if (tabId === 'dashboard') initDash();
            else if (tabId === 'tiering')   calcT();
            else if (tabId === 'profit')    calcP();
            else if (tabId === 'promo')     cbRender();
        }
    }
    if (id === 'branch-view') bcCalculate();
}

// ── Chart color helper — reads current theme ──
function getChartColors() {
    const isDark = document.body.getAttribute('data-theme') === 'dark';
    return {
        neutral: isDark ? '#3d4149' : '#e9ecef',
        dark:    isDark ? '#f1f3f5' : '#1a1a2e',
        gridColor: isDark ? 'rgba(255,255,255,0.06)' : 'rgba(0,0,0,0.06)',
        tickColor: isDark ? '#adb5bd' : '#6c757d',
    };
}
function updateChartTheme() {
    const c = getChartColors();
    if (funnelChart) {
        funnelChart.data.datasets[0].backgroundColor = [c.neutral, '#d61e30', c.dark];
        funnelChart.options.scales.x = funnelChart.options.scales.x || {};
        funnelChart.options.scales.x.ticks = { color: c.tickColor };
        funnelChart.options.scales.x.grid  = { color: c.gridColor };
        funnelChart.options.scales.y = funnelChart.options.scales.y || {};
        funnelChart.options.scales.y.ticks = { color: c.tickColor };
        funnelChart.update();
    }
    if (growthChart) {
        growthChart.data.datasets[0].backgroundColor = [c.dark, '#d61e30'];
        growthChart.options.scales.y.ticks.color = c.tickColor;
        growthChart.options.scales.y.grid = { color: c.gridColor };
        growthChart.options.plugins.datalabels.color = c.tickColor;
        growthChart.update();
    }
}

function toggleTheme() {
    const current = document.body.getAttribute('data-theme') || document.documentElement.getAttribute('data-theme') || 'light';
    const t = current === 'light' ? 'dark' : 'light';
    document.body.setAttribute('data-theme', t);
    document.documentElement.setAttribute('data-theme', t);
    updateChartTheme();
    saveData();
}

// ── JSON Backup / Restore ──
function exportJSON() {
    try {
        const revData = localStorage.getItem(LS_KEY_REV);
        const sobunData = localStorage.getItem(LS_KEY_SOBUN);
        const backup = {
            _meta: { version: 1, exported: new Date().toISOString(), app: APP_NAME },
            revenue: revData ? JSON.parse(revData) : null,
            sobun: sobunData ? JSON.parse(sobunData) : null,
            branch: (() => { try { const b = localStorage.getItem(LS_KEY_BRANCH); return b ? JSON.parse(b) : null; } catch(e) { return null; } })(),
            targets: (() => { try { const t = localStorage.getItem(LS_KEY_TARGETS); return t ? JSON.parse(t) : null; } catch(e) { return null; } })(),
            history: (() => { try { const h = localStorage.getItem(LS_KEY_HISTORY); return h ? JSON.parse(h) : null; } catch(e) { return null; } })()
        };
        const blob = new Blob([JSON.stringify(backup, null, 2)], { type: 'application/json' });
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        const d = new Date(); 
        a.download = `taobun_backup_${d.getFullYear()}${String(d.getMonth()+1).padStart(2,'0')}${String(d.getDate()).padStart(2,'0')}.json`;
        a.click();
        URL.revokeObjectURL(a.href);
        xlToast('✅ Backup JSON berhasil diunduh!');
    } catch(e) { xlToast('⚠️ Gagal export backup: ' + e.message); }
}

function importFromJSON(input) {
    const file = input.files[0];
    if (!file) return;
    showImportLoading('Memverifikasi backup...');
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const backup = JSON.parse(e.target.result);
            if (!backup._meta || backup._meta.app !== APP_NAME) {
                hideImportLoading();
                xlToast('⚠️ File bukan backup TAOBUN yang valid.');
                return;
            }
            showImportLoading('Mengembalikan data...');
            if (backup.revenue) localStorage.setItem(LS_KEY_REV, JSON.stringify(backup.revenue));
            if (backup.sobun) localStorage.setItem(LS_KEY_SOBUN, JSON.stringify(backup.sobun));
            if (backup.branch) localStorage.setItem(LS_KEY_BRANCH, JSON.stringify(backup.branch));
            if (backup.targets) localStorage.setItem(LS_KEY_TARGETS, JSON.stringify(backup.targets));
            if (backup.history) localStorage.setItem(LS_KEY_HISTORY, JSON.stringify(backup.history));
            xlToast('✅ Data berhasil di-restore! Memuat ulang...');
            setTimeout(() => location.reload(), 1800);
        } catch(e) {
            hideImportLoading();
            xlToast('⚠️ File JSON tidak valid: ' + e.message);
        }
    };
    reader.onerror = () => { hideImportLoading(); xlToast('⚠️ Gagal membaca file.'); };
    reader.readAsText(file);
    input.value = '';
}

function toggleDropdown(id) {
    const el = document.getElementById(id);
    const wasActive = el.classList.contains('active');
    document.querySelectorAll('.custom-select-container').forEach(c => c.classList.remove('active'));
    if (!wasActive) el.classList.add('active');
}

function selectOption(type, value, label) {
    if (type === 'Month') { 
        currentMonth = value; 
        document.querySelectorAll('[id^="selectedMonth"]').forEach(el => el.textContent = label);
    } else { 
        currentYear = value; 
        document.querySelectorAll('[id^="selectedYear"]').forEach(el => el.textContent = value);
    }
    document.querySelectorAll('.custom-select-container').forEach(c => c.classList.remove('active'));
    calculate();
    calculateLeads();
    updatePeriodBadge();
}

function updatePeriodBadge() {
    const mEl = document.getElementById('selectedMonthBranch');
    const yEl = document.getElementById('selectedYearBranch');
    const badge = document.getElementById('branch-period-text');
    if (badge && mEl && yEl) badge.textContent = `${mEl.textContent} ${yEl.textContent}`;
}

function handleInput(el) {
  let cursor = el.selectionStart, oldLen = el.value.length;
  el.value = formatNumber(el.value);
  el.setSelectionRange(cursor + (el.value.length - oldLen), cursor + (el.value.length - oldLen));
  clearValidation(el);
  calculate();
}

function handleLeadInput(el) {
  el.value = formatNumber(el.value);
  clearValidation(el);
  calculateLeads();
}
function syncSlider(val) { document.getElementById('growthSlider').value = cleanNumber(val); calculate(); }
function syncInput(val) { document.getElementById('optPercent').value = formatNumber(val); calculate(); }

// ── Micro-animation helper ──
function animateEl(id) {
  const el = document.getElementById(id);
  if (!el) return;
  el.classList.remove('value-pop');
  void el.offsetWidth; // reflow to restart animation
  el.classList.add('value-pop');
}

function calculate() {
  const leads = cleanNumber(document.getElementById('leads').value);
  const convTarget = cleanNumber(document.getElementById('convTargetInput').value);
  const cust = cleanNumber(document.getElementById('totalCustInput').value);
  const ret = cleanNumber(document.getElementById('retentionCust').value);
  const basket = cleanNumber(document.getElementById('basketSize').value);
  const avgTx = cleanNumber(document.getElementById('avgTx').value);
  const opt = cleanNumber(document.getElementById('optPercent').value);
  
  const actualConv = leads > 0 ? (cust / leads * 100).toFixed(1) : 0;
  const newCust = Math.max(0, cust - ret);
  
  document.getElementById('actualConvDisplay').textContent = actualConv + '%';
  document.getElementById('newCustDisplay').textContent = formatNumber(Math.round(newCust));

  const currentRev = cust * avgTx * basket;
  const optRev = currentRev * Math.pow(1 + (opt/100), 3);
  
  document.getElementById('totalRevenue').textContent = 'Rp ' + currentRev.toLocaleString('id-ID');
  document.getElementById('projectedRevenueDisplay').textContent = 'Rp ' + optRev.toLocaleString('id-ID');
  animateEl('totalRevenue');
  animateEl('projectedRevenueDisplay');

  const panel = document.getElementById('insightPanel');
  if (leads > 0 && parseFloat(actualConv) < convTarget) {
      panel.innerHTML = `<b>💡 Business Insight:</b> Konversi (${actualConv}%) di bawah target (${convTarget}%).`;
      panel.style.display = 'block';
  } else { panel.style.display = 'none'; }

  if (funnelChart) { funnelChart.data.datasets[0].data = [leads, cust, cust * avgTx]; funnelChart.update(); }
  if (growthChart) { growthChart.data.datasets[0].data = [currentRev, optRev]; growthChart.update(); }
  saveData();
}

function calculateLeads() {
    const reach = cleanNumber(document.getElementById('reach').value);
    const visits = cleanNumber(document.getElementById('visits').value);
    const links = cleanNumber(document.getElementById('links').value);
    const orders = cleanNumber(document.getElementById('orders').value);
    const payments = cleanNumber(document.getElementById('payments').value);
    const res = document.getElementById('leads-results');

    if (reach > 0) {
        const r2v = ((visits/reach)*100).toFixed(2);
        const v2l = visits > 0 ? ((links/visits)*100).toFixed(2) : 0;
        const l2o = links > 0 ? ((orders/links)*100).toFixed(2) : 0;
        const o2p = orders > 0 ? ((payments/orders)*100).toFixed(2) : 0;
        const ovr = ((payments/reach)*100).toFixed(2);

        res.innerHTML = `
            <div class="funnel-item"><span>Reach → Visit</span><span class="result-val">${r2v}%</span></div>
            <div class="funnel-item"><span>Visit → Link</span><span class="result-val">${v2l}%</span></div>
            <div class="funnel-item"><span>Link → Order</span><span class="result-val">${l2o}%</span></div>
            <div class="funnel-item"><span>Order → Pay</span><span class="result-val">${o2p}%</span></div>
            <div class="overall-row"><small>OVERALL CONV.</small><div style="font-size:32px; font-weight:900;">${ovr}%</div></div>`;
    } else {
        res.innerHTML = `<p style="text-align: center; color: var(--muted);">Lengkapi data di atas.</p>`;
    }
}

function exportToExcel() {
  loadXlsx(function() { _doExportToExcel(); });
}
function _doExportToExcel() {
    // ── Helpers ──
    const rp = v => { const n = typeof v === 'string' ? cleanNumber(v) : v; return isNaN(n)||n===0 ? 'Rp 0' : 'Rp ' + Math.round(n).toLocaleString('id-ID'); };
    const num = v => typeof v === 'string' ? cleanNumber(v) : (v||0);
    const pct = v => { const n = typeof v === 'string' ? parseFloat(v) : v; return isNaN(n) ? '0%' : n.toFixed(1) + '%'; };

    const mLabel = document.getElementById('selectedMonthRev')?.textContent || currentMonth;
    const yr = currentYear;
    const wb = XLSX.utils.book_new();

    // ─────────────────────────────────
    // Sheet 1: Revenue Dashboard
    // ─────────────────────────────────
    const leads       = num(document.getElementById('leads').value);
    const convTarget  = num(document.getElementById('convTargetInput').value);
    const cust        = num(document.getElementById('totalCustInput').value);
    const ret         = num(document.getElementById('retentionCust').value);
    const basket      = num(document.getElementById('basketSize').value);
    const avgTx       = num(document.getElementById('avgTx').value);
    const opt         = num(document.getElementById('optPercent').value);
    const totalRev    = cust * avgTx * basket;
    const projRev     = totalRev * Math.pow(1 + opt/100, 3);
    const actConv     = leads > 0 ? (cust/leads*100) : 0;
    const newCust     = Math.max(0, cust - ret);

    const revRows = [
        ['REVENUE DASHBOARD — TAOBUN', `${mLabel} ${yr}`],
        [],
        ['METRIK', 'NILAI', 'KETERANGAN'],
        ['Potential Leads', leads, 'Jumlah prospek/reach'],
        ['Target Conversion Rate', pct(convTarget), 'Target % leads jadi customer'],
        ['Actual Conversion Rate', pct(actConv), 'Aktual: Active Cust / Leads'],
        ['Active Customers', cust, 'Total pelanggan aktif bulan ini'],
        ['New Customers', newCust, 'Cust baru = Active - Retained'],
        ['Retained Customers', ret, 'Customer yang kembali'],
        ['Basket Size (Avg Belanja)', rp(basket), 'Rata-rata nilai 1 transaksi'],
        ['Avg Transaksi / Bulan', avgTx, 'Frekuensi beli per customer/bulan'],
        [],
        ['HASIL KALKULASI', '', ''],
        ['Total Revenue (Aktual)', rp(totalRev), 'Active Cust × Avg Tx × Basket'],
        ['Projected Revenue (Optimasi)', rp(projRev), `Simulasi pertumbuhan +${opt}%`],
        ['Selisih Proyeksi vs Aktual', rp(projRev - totalRev), 'Potensi kenaikan revenue'],
    ];

    // ─────────────────────────────────
    // Sheet 2: Lead Funnel
    // ─────────────────────────────────
    const reach    = num(document.getElementById('reach').value);
    const visits   = num(document.getElementById('visits').value);
    const links    = num(document.getElementById('links').value);
    const orders   = num(document.getElementById('orders').value);
    const payments = num(document.getElementById('payments').value);
    const r2v = reach  > 0 ? (visits/reach*100)   : 0;
    const v2l = visits > 0 ? (links/visits*100)    : 0;
    const l2o = links  > 0 ? (orders/links*100)    : 0;
    const o2p = orders > 0 ? (payments/orders*100) : 0;
    const ovr = reach  > 0 ? (payments/reach*100)  : 0;

    const leadRows = [
        ['LEAD CONVERSION FUNNEL — TAOBUN', `${mLabel} ${yr}`],
        [],
        ['STAGE', 'VOLUME', 'CONV. RATE KE STAGE BERIKUTNYA'],
        ['Reach / Impressions', reach, '—'],
        ['Profile Visits', visits, pct(r2v) + ' (Reach → Visit)'],
        ['External Link Clicks', links, pct(v2l) + ' (Visit → Link)'],
        ['ESB Order Clicks', orders, pct(l2o) + ' (Link → Order)'],
        ['Success Payments', payments, pct(o2p) + ' (Order → Pay)'],
        [],
        ['OVERALL CONVERSION', pct(ovr), 'Payments / Reach'],
    ];

    // ─────────────────────────────────
    // Sheet 3: Branch Comparison
    // ─────────────────────────────────
    const bGet = id => num(document.getElementById(id)?.value || '0');
    const branches = [
        { label: 'Taobun Raya',    rev: bGet('br-raya-rev'),  cust: bGet('br-raya-cust'),  tx: bGet('br-raya-tx'),  basket: bGet('br-raya-basket'),  newC: bGet('br-raya-new'),  cogs: bGet('br-raya-cogs')  },
        { label: 'Taobun Kemboja', rev: bGet('br-kemb-rev'),  cust: bGet('br-kemb-cust'),  tx: bGet('br-kemb-tx'),  basket: bGet('br-kemb-basket'),  newC: bGet('br-kemb-new'),  cogs: bGet('br-kemb-cogs')  },
        { label: 'Taobun Perdos',  rev: bGet('br-perd-rev'),  cust: bGet('br-perd-cust'),  tx: bGet('br-perd-tx'),  basket: bGet('br-perd-basket'),  newC: bGet('br-perd-new'),  cogs: bGet('br-perd-cogs')  },
    ];

    const branchRows = [
        ['BRANCH COMPARISON — TAOBUN', `${mLabel} ${yr}`],
        [],
        ['METRIK', 'Taobun Raya', 'Taobun Kemboja', 'Taobun Perdos', 'TOTAL / RATA-RATA'],
    ];
    const metrics = [
        { label: 'Revenue (Rp)',         key: 'rev',    fmt: rp  },
        { label: 'Active Customers',     key: 'cust',   fmt: v=>v },
        { label: 'Transaksi',            key: 'tx',     fmt: v=>v },
        { label: 'Avg Basket (Rp)',      key: 'basket', fmt: rp  },
        { label: 'New Customers',        key: 'newC',   fmt: v=>v },
        { label: 'COGS / HPP (Rp)',      key: 'cogs',   fmt: rp  },
        { label: 'Gross Margin (Rp)',    key: null,     fmt: null },
        { label: 'Gross Margin (%)',     key: null,     fmt: null },
    ];
    metrics.forEach(m => {
        if (m.key) {
            const vals = branches.map(b => b[m.key]);
            const total = vals.reduce((a,c)=>a+c,0);
            branchRows.push([m.label, m.fmt(vals[0]), m.fmt(vals[1]), m.fmt(vals[2]), m.fmt(total)]);
        } else if (m.label === 'Gross Margin (Rp)') {
            const gms = branches.map(b => b.rev - b.cogs);
            branchRows.push([m.label, rp(gms[0]), rp(gms[1]), rp(gms[2]), rp(gms.reduce((a,c)=>a+c,0))]);
        } else {
            const gmps = branches.map(b => b.rev > 0 ? (b.rev-b.cogs)/b.rev*100 : 0);
            branchRows.push([m.label, pct(gmps[0]), pct(gmps[1]), pct(gmps[2]), pct(gmps.reduce((a,c)=>a+c,0)/3)]);
        }
    });

    // ─────────────────────────────────
    // Build & style sheets
    // ─────────────────────────────────
    const sheets = [
        { name: 'Revenue Dashboard', rows: revRows },
        { name: 'Lead Funnel',       rows: leadRows },
        { name: 'Branch Comparison', rows: branchRows },
    ];

    // Column widths per sheet
    const colWidths = [
        [{ wch: 34 }, { wch: 28 }, { wch: 36 }],
        [{ wch: 32 }, { wch: 18 }, { wch: 38 }],
        [{ wch: 26 }, { wch: 22 }, { wch: 22 }, { wch: 22 }, { wch: 22 }],
    ];

    sheets.forEach((s, si) => {
        const ws = XLSX.utils.aoa_to_sheet(s.rows);
        ws['!cols'] = colWidths[si];
        XLSX.utils.book_append_sheet(wb, ws, s.name);
    });

    const d = new Date();
    const stamp = `${d.getFullYear()}${String(d.getMonth()+1).padStart(2,'0')}${String(d.getDate()).padStart(2,'0')}`;
    XLSX.writeFile(wb, `Report_Taobun_${currentMonth}_${currentYear}_${stamp}.xlsx`);
    xlToast('✅ Report Excel berhasil diekspor — 3 sheet.');
}

function importFromExcel(input) {
  loadXlsx(function() { _doImportFromExcel(input); });
}
function _doImportFromExcel(input) {
    if (!file) return;
    if (!file.name.match(/\.xlsx?$/i)) { xlToast('⚠️ File harus berformat .xlsx atau .xls'); input.value = ''; return; }
    showImportLoading('Membaca file Excel...');
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            showImportLoading('Parsing data...');
            const data = new Uint8Array(e.target.result);
            const wb = XLSX.read(data, { type: 'array' });
            const ws = wb.Sheets[wb.SheetNames[0]];
            const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

            // Field map: Excel label → { id, type, label }
            // type: 'number' (plain), 'rp' (rupiah formatted), 'pct' (percentage)
            const fieldMap = {
                'Potential Leads':      { id: 'leads',           type: 'number', label: 'Potential Leads' },
                'Active Customers':     { id: 'totalCustInput',  type: 'number', label: 'Active Customers' },
                'Basket Size (Avg Belanja)': { id: 'basketSize', type: 'number', label: 'Basket Size' },
                'Basket Size':          { id: 'basketSize',      type: 'number', label: 'Basket Size' },
                'Avg Transaksi / Bulan':{ id: 'avgTx',           type: 'number', label: 'Avg Tx/Bulan' },
                'Avg Tx':               { id: 'avgTx',           type: 'number', label: 'Avg Tx' },
                'Avg Tx / Month':       { id: 'avgTx',           type: 'number', label: 'Avg Tx' },
                'Target Conversion Rate':{ id: 'convTargetInput',type: 'pct',    label: 'Target Rate' },
                'Target Rate':          { id: 'convTargetInput', type: 'pct',    label: 'Target Rate' },
                'Retained Customers':   { id: 'retentionCust',   type: 'number', label: 'Retained Cust' },
                'Retention':            { id: 'retentionCust',   type: 'number', label: 'Retained Cust' },
                'Growth Ambition':      { id: 'optPercent',      type: 'pct',    label: 'Growth Ambition' },
                // Lead Funnel fields
                'Reach / Impressions':  { id: 'reach',           type: 'number', label: 'Reach' },
                'Profile Visits':       { id: 'visits',          type: 'number', label: 'Profile Visits' },
                'External Link Clicks': { id: 'links',           type: 'number', label: 'Link Clicks' },
                'ESB Order Clicks':     { id: 'orders',          type: 'number', label: 'Order Clicks' },
                'Success Payments':     { id: 'payments',        type: 'number', label: 'Payments' },
            };

            const changes = [], skipped = [];
            rows.forEach(row => {
                const label = String(row[0] || '').trim();
                const rawVal = row[1];
                if (!label || rawVal === '' || rawVal == null) return;
                const field = fieldMap[label];
                if (!field) return;
                const el = document.getElementById(field.id);
                if (!el) return;

                let newNum;
                if (field.type === 'pct') {
                    // Accept "12.5%", "12.5", or 0.125
                    const str = String(rawVal).replace('%','').trim();
                    const parsed = parseFloat(str);
                    // If stored as decimal (0–1 range), convert to percent
                    newNum = (!isNaN(parsed) && parsed > 0 && parsed <= 1 && !str.includes('.')) ? parsed * 100 : parsed;
                } else {
                    // Strip Rp, dots, commas
                    const str = String(rawVal).replace(/Rp\s?/gi,'').replace(/\./g,'').replace(/,/g,'.').trim();
                    newNum = parseFloat(str);
                }

                if (isNaN(newNum)) { skipped.push(field.label); return; }
                newNum = Math.round(newNum);

                const oldNum = cleanNumber(el.value);
                if (newNum === oldNum) return; // tidak berubah

                changes.push({ field: field.label, oldVal: oldNum, newVal: newNum, el, type: field.type });
            });

            hideImportLoading();

            if (changes.length === 0 && skipped.length === 0) {
                xlToast('ℹ️ Tidak ada perubahan yang terdeteksi.');
                input.value = '';
                return;
            }

            if (changes.length === 0) {
                xlToast(`⚠️ Tidak ada field yang cocok atau berubah. (${skipped.length} field dilewati)`);
                input.value = '';
                return;
            }

            // Apply changes
            const rpFmtLocal = v => 'Rp ' + v.toLocaleString('id-ID');
            changes.forEach(c => {
                c.el.value = formatNumber(String(c.newVal));
                clearValidation(c.el);
            });

            // Build human-readable change log
            const changeLog = changes.map(c => {
                const fmt = v => c.type === 'pct' ? v + '%' : c.type === 'rp' ? rpFmtLocal(v) : v.toLocaleString('id-ID');
                return `${c.field}: ${fmt(c.oldVal)} → ${fmt(c.newVal)}`;
            }).join('\n');

            calculate();
            calculateLeads();
            saveData();

            // Show rich summary toast
            const skipNote = skipped.length ? `\n(${skipped.length} field dilewati)` : '';
            xlToast(`✅ ${changes.length} field diperbarui dari Excel!${skipNote}`);

            // Console log detail for debugging
            console.log('[Import Excel - Revenue] Perubahan terdeteksi:\n' + changeLog);

        } catch (err) {
            hideImportLoading();
            xlToast('⚠️ Gagal membaca file: ' + err.message);
        }
        input.value = '';
    };
    reader.onerror = () => { hideImportLoading(); xlToast('⚠️ Gagal membaca file.'); };
    reader.readAsArrayBuffer(file);
}

function initCharts() {
  if (typeof Chart === 'undefined') return;
  const c = getChartColors();
  funnelChart = new Chart(document.getElementById('funnelChart'), {
    type: 'bar',
    data: { labels: ['Leads', 'Users', 'Orders'], datasets: [{ data: [0,0,0], backgroundColor: [c.neutral, '#d61e30', c.dark], borderRadius: 10 }] },
    options: {
        indexAxis: 'y', responsive: true, maintainAspectRatio: false,
        plugins: { legend: {display: false}, datalabels: { display: false } },
        scales: {
            x: { ticks: { color: c.tickColor }, grid: { color: c.gridColor } },
            y: { ticks: { color: c.tickColor }, grid: { color: c.gridColor } }
        }
    }
  });
  
  growthChart = new Chart(document.getElementById('growthChart'), {
    type: 'bar',
    data: { labels: ['Current', 'Target'], datasets: [{ data: [0,0], backgroundColor: [c.dark, '#d61e30'], borderRadius: 12 }] },
    options: { 
        responsive: true, maintainAspectRatio: false, layout: { padding: { top: 30 } }, 
        plugins: {
            legend: {display: false},
            datalabels: { anchor: 'end', align: 'top', offset: 5, color: c.tickColor, font: {weight: '800'}, formatter: (v) => 'Rp' + formatShortNumber(v) }
        },
        scales: {
            y: {
                beginAtZero: true,
                ticks: { color: c.tickColor, callback: function(value) { return 'Rp' + formatShortNumber(value); } },
                grid: { color: c.gridColor }
            }
        }
    }
  });
}

function saveData() {
    const data = { 
        l: document.getElementById('leads').value, 
        c: document.getElementById('convTargetInput').value, 
        t: document.getElementById('totalCustInput').value, 
        r: document.getElementById('retentionCust').value, 
        b: document.getElementById('basketSize').value, 
        tx: document.getElementById('avgTx').value, 
        o: document.getElementById('optPercent').value,
        lead_reach: document.getElementById('reach').value,
        lead_visits: document.getElementById('visits').value,
        lead_links: document.getElementById('links').value,
        lead_orders: document.getElementById('orders').value,
        lead_payments: document.getElementById('payments').value,
        month: currentMonth,
        year: currentYear,
        theme: document.body.getAttribute('data-theme') 
    };
    localStorage.setItem(LS_KEY_REV, JSON.stringify(data));
}

function loadData() {
    const saved = localStorage.getItem(LS_KEY_REV);
    if (saved) {
        const d = JSON.parse(saved);
        document.getElementById('leads').value = d.l || "0";
        document.getElementById('convTargetInput').value = d.c || "0";
        document.getElementById('totalCustInput').value = d.t || "0";
        document.getElementById('retentionCust').value = d.r || "0";
        document.getElementById('basketSize').value = d.b || "0";
        document.getElementById('avgTx').value = d.tx || "0";
        document.getElementById('optPercent').value = d.o || "0";
        document.getElementById('reach').value = d.lead_reach || "0";
        document.getElementById('visits').value = d.lead_visits || "0";
        document.getElementById('links').value = d.lead_links || "0";
        document.getElementById('orders').value = d.lead_orders || "0";
        document.getElementById('payments').value = d.lead_payments || "0";
        document.getElementById('growthSlider').value = cleanNumber(d.o) || 0;
        // Restore month/year selection
        if (d.month) {
            currentMonth = d.month;
            const mLabel = months.find(m => m.v === d.month)?.l || d.month;
            document.querySelectorAll('[id^="selectedMonth"]').forEach(el => el.textContent = mLabel);
        }
        if (d.year) {
            currentYear = d.year;
            document.querySelectorAll('[id^="selectedYear"]').forEach(el => el.textContent = d.year);
        }
        if (d.theme) {
            document.body.setAttribute('data-theme', d.theme);
            document.documentElement.setAttribute('data-theme', d.theme);
        }
        calculate();
        calculateLeads();
    }
}

function resetMetrics() {
    showConfirm({
        icon: '🗑️', title: 'Reset Revenue Data?',
        msg: 'Semua data di Revenue Dashboard akan dikosongkan.',
        okLabel: 'Reset'
    }, function() {
        const ids = ['leads', 'convTargetInput', 'totalCustInput', 'retentionCust', 'basketSize', 'avgTx', 'optPercent'];
        ids.forEach(id => { const el = document.getElementById(id); if (el) el.value = '0'; });
        document.getElementById('growthSlider').value = 0;
        document.querySelectorAll('.nav-input').forEach(el => clearValidation(el));
        saveData();
        calculate();
        xlToast('✅ Data Revenue berhasil direset.');
    });
}

function resetLeadMetrics() {
    showConfirm({
        icon: '🗑️', title: 'Reset Funnel Data?',
        msg: 'Semua data funnel di Lead Management akan dikosongkan.',
        okLabel: 'Reset'
    }, function() {
        const ids = ['reach', 'visits', 'links', 'orders', 'payments'];
        ids.forEach(id => { const el = document.getElementById(id); if (el) el.value = '0'; });
        document.querySelectorAll('.lead-field').forEach(el => clearValidation(el));
        saveData();
        calculateLeads();
        xlToast('✅ Data Funnel berhasil direset.');
    });
}

/* =============================================
   SOBAT TAOBUN — JAVASCRIPT
   ============================================= */
const SB_P = [
  {n:'ES KOPI KEMBOJA',hpp:9352,hj:25000,gm:15648,gmp:.62592,par:'Non-Pareto',red:'BISA',poin:470,mt:30000},
  {n:'HOT KOPI KEMBOJA',hpp:9138,hj:23000,gm:13862,gmp:.6027,par:'Non-Pareto',red:'BISA',poin:460,mt:29000},
  {n:'ICE AMERICANO',hpp:5992,hj:22000,gm:16008,gmp:.7276,par:'Pareto A',red:'BISA',poin:300,mt:19000},
  {n:'HOT AMERICANO',hpp:6094,hj:20000,gm:13906,gmp:.6953,par:'Non-Pareto',red:'BISA',poin:305,mt:20000},
  {n:'ICE KOPI KEMUNING',hpp:6994,hj:23000,gm:16006,gmp:.6959,par:'Non-Pareto',red:'BISA',poin:350,mt:23000},
  {n:'HOT KEMUNING',hpp:7168,hj:20000,gm:12832,gmp:.6416,par:'Non-Pareto',red:'BISA',poin:360,mt:23000},
  {n:'ICE COFFEE LATTE',hpp:5630,hj:23000,gm:17370,gmp:.7552,par:'Non-Pareto',red:'BISA',poin:285,mt:18000},
  {n:'HOT COFFEE LATTE',hpp:7693,hj:20000,gm:12307,gmp:.6154,par:'Non-Pareto',red:'BISA',poin:385,mt:25000},
  {n:'ICE AMERICANO AREN',hpp:7862,hj:22000,gm:14138,gmp:.6426,par:'Non-Pareto',red:'BISA',poin:395,mt:25000},
  {n:'CARAMEL MACCHIATO',hpp:10072,hj:28000,gm:17928,gmp:.6403,par:'Non-Pareto',red:'BISA',poin:505,mt:32000},
  {n:'ICE HAZELNUT LATTE',hpp:8545,hj:28000,gm:19455,gmp:.6948,par:'Non-Pareto',red:'BISA',poin:430,mt:28000},
  {n:'HOT HAZELNUT LATTE',hpp:9751,hj:28000,gm:18249,gmp:.6518,par:'Non-Pareto',red:'BISA',poin:490,mt:31000},
  {n:'KOPI PANDAN SRIKAYA',hpp:9096,hj:28000,gm:18904,gmp:.6751,par:'Non-Pareto',red:'BISA',poin:455,mt:29000},
  {n:'ICE AMERICANO LEMONERO',hpp:7932,hj:25000,gm:17068,gmp:.6827,par:'Non-Pareto',red:'BISA',poin:400,mt:26000},
  {n:'ICE LEMON TEA',hpp:4473,hj:22000,gm:17527,gmp:.7967,par:'Non-Pareto',red:'BISA',poin:225,mt:15000},
  {n:'HOT LEMON TEA',hpp:3241,hj:19000,gm:15759,gmp:.8294,par:'Non-Pareto',red:'BISA',poin:165,mt:11000},
  {n:'LYCHEE TEA',hpp:6066,hj:23000,gm:16934,gmp:.7363,par:'Non-Pareto',red:'BISA',poin:305,mt:20000},
  {n:'HOT TEH TARIK',hpp:4232,hj:20000,gm:15768,gmp:.7884,par:'Non-Pareto',red:'BISA',poin:215,mt:14000},
  {n:'ICE TEH TARIK',hpp:3770,hj:22000,gm:18230,gmp:.8286,par:'Non-Pareto',red:'BISA',poin:190,mt:12000},
  {n:'ICE CHOCO',hpp:6988,hj:23000,gm:16012,gmp:.6962,par:'Pareto A',red:'BISA',poin:350,mt:23000},
  {n:'ICE RED VELVET',hpp:9969,hj:28000,gm:18031,gmp:.644,par:'Non-Pareto',red:'BISA',poin:500,mt:32000},
  {n:'ICE MATCHA',hpp:5718,hj:28000,gm:22282,gmp:.7958,par:'Pareto B',red:'BISA',poin:290,mt:19000},
  {n:'STRAWBERRY MATCHA',hpp:7892,hj:28000,gm:20108,gmp:.7181,par:'Pareto B',red:'BISA',poin:395,mt:25000},
  {n:'HOT MATCHA',hpp:9679,hj:28000,gm:18321,gmp:.6543,par:'Non-Pareto',red:'BISA',poin:485,mt:31000},
  {n:'ICE REGAL',hpp:7219,hj:23000,gm:15781,gmp:.6861,par:'Non-Pareto',red:'BISA',poin:365,mt:23000},
  {n:'ICE PISANG COKLAT',hpp:7264,hj:25000,gm:17736,gmp:.7094,par:'Non-Pareto',red:'BISA',poin:365,mt:23000},
  {n:'COTTON CANDY',hpp:6521,hj:23000,gm:16479,gmp:.7165,par:'Non-Pareto',red:'BISA',poin:330,mt:21000},
  {n:'PANDAN SRIKAYA',hpp:8732,hj:22000,gm:13268,gmp:.6031,par:'Non-Pareto',red:'BISA',poin:440,mt:28000},
  {n:'PIJAR MATAHARI',hpp:8667,hj:28000,gm:19333,gmp:.6905,par:'Non-Pareto',red:'BISA',poin:435,mt:28000},
  {n:'MANGGUO',hpp:7337,hj:22000,gm:14663,gmp:.6665,par:'Non-Pareto',red:'BISA',poin:370,mt:24000},
  {n:'PINEAPLE EXPRESS',hpp:6198,hj:23000,gm:16802,gmp:.7305,par:'Non-Pareto',red:'BISA',poin:310,mt:20000},
  {n:'ROTI KUKUS SRIKAYA',hpp:2530,hj:6500,gm:3970,gmp:.6108,par:'Pareto A',red:'BISA',poin:130,mt:9000},
  {n:'ROTI KUKUS COKLAT',hpp:2415,hj:6500,gm:4085,gmp:.6285,par:'Pareto A',red:'BISA',poin:125,mt:8000},
  {n:'ROTI KUKUS KEJU',hpp:2530,hj:7000,gm:4470,gmp:.6386,par:'Pareto A',red:'BISA',poin:130,mt:9000},
  {n:'ROTI GORENG COKLAT',hpp:2415,hj:6500,gm:4085,gmp:.6285,par:'Pareto B',red:'BISA',poin:125,mt:8000},
  {n:'ROTI GORENG KEJU',hpp:2530,hj:7000,gm:4470,gmp:.6386,par:'Non-Pareto',red:'BISA',poin:130,mt:9000},
  {n:'BIGBUN BLACK PAPPER',hpp:9810,hj:25000,gm:15190,gmp:.6076,par:'Non-Pareto',red:'BISA',poin:495,mt:32000},
  {n:'INDOMIE GORENG SINGLE',hpp:5762,hj:18000,gm:12238,gmp:.6799,par:'Non-Pareto',red:'BISA',poin:290,mt:19000},
  {n:'INDOMIE SOTO SINGLE',hpp:5675,hj:18000,gm:12325,gmp:.6847,par:'Non-Pareto',red:'BISA',poin:285,mt:18000},
  {n:'INDOMIE KALDU SINGLE',hpp:5373,hj:18000,gm:12627,gmp:.7015,par:'Non-Pareto',red:'BISA',poin:270,mt:17000},
  {n:'INDOMIE KALDU DOUBLE',hpp:8378,hj:22000,gm:13622,gmp:.6192,par:'Non-Pareto',red:'BISA',poin:420,mt:27000},
  {n:'INDOMIE COTO SINGLE',hpp:5963,hj:18000,gm:12037,gmp:.6687,par:'Non-Pareto',red:'BISA',poin:300,mt:19000},
  {n:'KOKAM SELITER',hpp:61507,hj:100000,gm:38493,gmp:.38493,par:'Non-Pareto',red:'TIDAK',poin:null,mt:null},
  {n:'TEH TARIK 1 LITER',hpp:26205,hj:85000,gm:58795,gmp:.6917,par:'Non-Pareto',red:'TIDAK',poin:null,mt:null},
  {n:'ROTI KUKUS AYAM',hpp:3565,hj:7500,gm:3935,gmp:.5247,par:'Pareto A',red:'TIDAK',poin:null,mt:null},
  {n:'ROTI KUKUS SAPI',hpp:4255,hj:8500,gm:4245,gmp:.4994,par:'Non-Pareto',red:'TIDAK',poin:null,mt:null},
  {n:'ROTI GORENG SRIKAYA',hpp:2645,hj:6500,gm:3855,gmp:.5931,par:'Non-Pareto',red:'TIDAK',poin:null,mt:null},
  {n:'ROTI GORENG AYAM',hpp:3565,hj:7500,gm:3935,gmp:.5247,par:'Pareto A',red:'TIDAK',poin:null,mt:null},
  {n:'ROTI GORENG SAPI',hpp:4255,hj:8500,gm:4245,gmp:.4994,par:'Pareto A',red:'TIDAK',poin:null,mt:null},
  {n:'NASI GORENG TAOBUN',hpp:13052,hj:30000,gm:16948,gmp:.5649,par:'Non-Pareto',red:'TIDAK',poin:null,mt:null},
  {n:'NASI GORENG HONGKONG',hpp:12229,hj:30000,gm:17771,gmp:.5924,par:'Pareto B',red:'TIDAK',poin:null,mt:null},
  {n:'NASI AYAM LADA HITAM',hpp:13098,hj:35000,gm:21902,gmp:.6258,par:'Non-Pareto',red:'TIDAK',poin:null,mt:null},
  {n:'CHICKEN KUNGPAO',hpp:16615,hj:35000,gm:18385,gmp:.5253,par:'Non-Pareto',red:'TIDAK',poin:null,mt:null},
  {n:'CHICKEN SALTED EGG',hpp:25924,hj:35000,gm:9076,gmp:.2593,par:'Non-Pareto',red:'TIDAK',poin:null,mt:null},
  {n:'NASI AYAM GEPREK',hpp:18838,hj:35000,gm:16162,gmp:.4618,par:'Non-Pareto',red:'TIDAK',poin:null,mt:null},
  {n:'NASI BEEF BLACKPAPPER',hpp:14631,hj:38000,gm:23369,gmp:.615,par:'Pareto B',red:'TIDAK',poin:null,mt:null},
  {n:'NASI BEEF TERIYAKI',hpp:15223,hj:38000,gm:22777,gmp:.5994,par:'Pareto B',red:'TIDAK',poin:null,mt:null},
  {n:'NASI AYAM SPICY',hpp:13025,hj:35000,gm:21975,gmp:.6279,par:'Non-Pareto',red:'TIDAK',poin:null,mt:null},
  {n:'NASI AYAM GORENG MENTEGA',hpp:14650,hj:35000,gm:20350,gmp:.5814,par:'Pareto B',red:'TIDAK',poin:null,mt:null},
  {n:'INDOMIE GORENG SALTED EGG',hpp:14532,hj:25000,gm:10468,gmp:.41872,par:'Non-Pareto',red:'TIDAK',poin:null,mt:null},
  {n:'INDOMIE GORENG KORNET',hpp:9337,hj:20000,gm:10663,gmp:.53315,par:'Non-Pareto',red:'TIDAK',poin:null,mt:null},
  {n:'INDOMIE GORENG DOUBLE',hpp:9154,hj:22000,gm:12846,gmp:.5839,par:'Non-Pareto',red:'TIDAK',poin:null,mt:null},
  {n:'INDOMIE SOTO DOUBLE',hpp:8982,hj:22000,gm:13018,gmp:.5917,par:'Non-Pareto',red:'TIDAK',poin:null,mt:null},
  {n:'INDOMIE COTO DOUBLE',hpp:9557,hj:22000,gm:12443,gmp:.5656,par:'Non-Pareto',red:'TIDAK',poin:null,mt:null},
];
// Use let so xlApply can mutate via .length=0 / .push()
let P = SB_P.slice();

const SB_TIERS = [
  {k:'bronze',lbl:'🥉 Bronze Bun',mult:1,css:'tb-br'},
  {k:'silver',lbl:'🥈 Silver Bun',mult:1.2,css:'tb-si'},
  {k:'gold',lbl:'🥇 Gold Bun',mult:1.5,css:'tb-go'},
  {k:'platinum',lbl:'💎 Platinum Bun',mult:2,css:'tb-pl'},
];

// Tier thresholds — dari data Excel SKEMA_SOBAT_TAOBUN
let TIER_TH = {
  bronze:   { naik: 3000000,   stay: 0        },
  silver:   { naik: 6000000,   stay: 4000000  },
  gold:     { naik: 10000000,  stay: 6500000  },
  platinum: { naik: 12000000,  stay: 12000000 },
};

let RW = [
  // ── BRONZE — Quick wins untuk akuisisi member baru ──
  {t:'bronze',n:'Selamat Datang, Bun! 🎉',item:'ROTI KUKUS SRIKAYA',hj:6500,hpp:2530,gm:.611,poin:130,ef:260000,rs:null,ket:'Quick win ~7x transaksi — cocok buat akuisisi member baru'},
  {t:'bronze',n:'Gigit Pertama Gratis',item:'ROTI KUKUS COKLAT',hj:6500,hpp:2415,gm:.628,poin:125,ef:250000,rs:null,ket:'Quick win ~7x transaksi — cocok buat akuisisi member baru'},
  {t:'bronze',n:'First Bite on Us',item:'ROTI KUKUS KEJU',hj:7000,hpp:2530,gm:.639,poin:130,ef:260000,rs:null,ket:'Quick win ~7x transaksi — cocok buat akuisisi member baru'},
  {t:'bronze',n:'Roti Goreng Perdana',item:'ROTI GORENG SRIKAYA',hj:6500,hpp:2645,gm:.593,poin:135,ef:270000,rs:null,ket:'Quick win ~8x transaksi — cocok buat akuisisi member baru'},
  {t:'bronze',n:'Ngopi Perdana Gratis ☕',item:'HOT AMERICANO',hj:20000,hpp:6094,gm:.695,poin:305,ef:610000,rs:null,ket:'Quick win ~17x transaksi — cocok buat akuisisi member baru'},
  {t:'bronze',n:'Teh Hangat Pertama',item:'HOT LEMON TEA',hj:19000,hpp:3241,gm:.829,poin:165,ef:330000,rs:null,ket:'Quick win ~9x transaksi — cocok buat akuisisi member baru'},
  {t:'bronze',n:'Mie Perdana Member Baru',item:'INDOMIE KALDU SINGLE',hj:18000,hpp:5373,gm:.702,poin:270,ef:540000,rs:null,ket:'Quick win ~15x transaksi — cocok buat akuisisi member baru'},
  // ── SILVER (STAY = Rp 4.000.000/thn) ──
  {t:'silver',n:'Silver Refresher ☕',item:'ICE AMERICANO',hj:22000,hpp:5992,gm:.728,poin:300,ef:600000,rs:.15,ket:'✅ Sangat worth it — 15% dari threshold STAY'},
  {t:'silver',n:'Minuman Favorit Gratis',item:'ICE LEMON TEA',hj:22000,hpp:4473,gm:.797,poin:225,ef:450000,rs:.11,ket:'✅ Sangat worth it — 11% dari threshold STAY'},
  {t:'silver',n:'Silver Chill Treat',item:'ICE TEH TARIK',hj:22000,hpp:3770,gm:.829,poin:190,ef:380000,rs:.10,ket:'✅ Sangat worth it — 10% dari threshold STAY'},
  {t:'silver',n:'Coklat Gratis Member Setia',item:'ICE CHOCO',hj:23000,hpp:6988,gm:.696,poin:350,ef:700000,rs:.17,ket:'👍 Wajar — 17% dari threshold STAY'},
  {t:'silver',n:'Silver Latte Day',item:'ICE COFFEE LATTE',hj:23000,hpp:5630,gm:.755,poin:285,ef:570000,rs:.14,ket:'✅ Sangat worth it — 14% dari threshold STAY'},
  {t:'silver',n:'Matcha Reward — Silver Only',item:'ICE MATCHA',hj:28000,hpp:5718,gm:.796,poin:290,ef:580000,rs:.14,ket:'✅ Sangat worth it — 14% dari threshold STAY'},
  {t:'silver',n:'Cotton Candy Vibes 🍬',item:'COTTON CANDY',hj:23000,hpp:6521,gm:.716,poin:330,ef:660000,rs:.17,ket:'👍 Wajar — 17% dari threshold STAY'},
  {t:'silver',n:'Lemon Tea Treat',item:'HOT LEMON TEA',hj:19000,hpp:3241,gm:.829,poin:165,ef:330000,rs:.08,ket:'✅ Sangat worth it — 8% dari threshold STAY'},
  {t:'silver',n:'Pineapple on Me 🍍',item:'PINEAPLE EXPRESS',hj:23000,hpp:6198,gm:.730,poin:310,ef:620000,rs:.15,ket:'✅ Sangat worth it — 15% dari threshold STAY'},
  {t:'silver',n:'Lychee Vibes — Silver',item:'LYCHEE TEA',hj:23000,hpp:6066,gm:.736,poin:305,ef:610000,rs:.15,ket:'✅ Sangat worth it — 15% dari threshold STAY'},
  {t:'silver',n:'Mie Double Gratis',item:'INDOMIE SOTO DOUBLE',hj:22000,hpp:8982,gm:.592,poin:450,ef:900000,rs:.23,ket:'👍 Wajar — 23% dari threshold STAY'},
  {t:'silver',n:'Mie Kornet — Silver Special',item:'INDOMIE GORENG KORNET',hj:20000,hpp:9337,gm:.533,poin:470,ef:940000,rs:.23,ket:'👍 Wajar — 23% dari threshold STAY'},
  // ── GOLD (STAY = Rp 6.500.000/thn) ──
  {t:'gold',n:'Gold Member Special ☕',item:'ES KOPI KEMBOJA',hj:25000,hpp:9352,gm:.626,poin:470,ef:940000,rs:.14,ket:'✅ Sangat worth it — 14% dari threshold STAY'},
  {t:'gold',n:'Signature Kopi Reward',item:'ICE KOPI KEMUNING',hj:23000,hpp:6994,gm:.696,poin:360,ef:720000,rs:.11,ket:'✅ Sangat worth it — 11% dari threshold STAY'},
  {t:'gold',n:'Aren Vibes — Gold Only',item:'ICE AMERICANO AREN',hj:22000,hpp:7862,gm:.643,poin:395,ef:790000,rs:.12,ket:'✅ Sangat worth it — 12% dari threshold STAY'},
  {t:'gold',n:'Matcha Premium Reward',item:'STRAWBERRY MATCHA',hj:28000,hpp:7892,gm:.718,poin:420,ef:840000,rs:.13,ket:'✅ Sangat worth it — 13% dari threshold STAY'},
  {t:'gold',n:'Hazelnut Latte on Us',item:'ICE HAZELNUT LATTE',hj:28000,hpp:8545,gm:.695,poin:455,ef:910000,rs:.14,ket:'✅ Sangat worth it — 14% dari threshold STAY'},
  {t:'gold',n:'Gold Pisang Coklat Treat',item:'ICE PISANG COKLAT',hj:25000,hpp:7264,gm:.709,poin:390,ef:780000,rs:.12,ket:'✅ Sangat worth it — 12% dari threshold STAY'},
  {t:'gold',n:'Regal Gold Member Drink',item:'ICE REGAL',hj:23000,hpp:7219,gm:.686,poin:375,ef:750000,rs:.12,ket:'✅ Sangat worth it — 12% dari threshold STAY'},
  {t:'gold',n:'Pijar Matahari — Gold Excl.',item:'PIJAR MATAHARI',hj:28000,hpp:8667,gm:.690,poin:490,ef:980000,rs:.15,ket:'✅ Sangat worth it — 15% dari threshold STAY'},
  {t:'gold',n:'Lemonero Gold Reward',item:'ICE AMERICANO LEMONERO',hj:25000,hpp:7932,gm:.683,poin:420,ef:840000,rs:.13,ket:'✅ Sangat worth it — 13% dari threshold STAY'},
  {t:'gold',n:'Pandan Srikaya — Gold Only',item:'PANDAN SRIKAYA',hj:22000,hpp:8732,gm:.603,poin:520,ef:1040000,rs:.16,ket:'👍 Wajar — 16% dari threshold STAY'},
  {t:'gold',n:'Nasi Hongkong — Gold Treat',item:'NASI GORENG HONGKONG',hj:30000,hpp:12229,gm:.592,poin:715,ef:1430000,rs:.22,ket:'👍 Wajar — 22% dari threshold STAY'},
  {t:'gold',n:'Mie Double Coto Gratis',item:'INDOMIE COTO DOUBLE',hj:22000,hpp:9557,gm:.566,poin:570,ef:1140000,rs:.18,ket:'👍 Wajar — 18% dari threshold STAY'},
  // ── PLATINUM (STAY = Rp 12.000.000/thn) ──
  {t:'platinum',n:'VIP Caramel Experience ✨',item:'CARAMEL MACCHIATO',hj:28000,hpp:10072,gm:.640,poin:510,ef:1020000,rs:.09,ket:'✅ Sangat worth it — 9% dari threshold STAY'},
  {t:'platinum',n:'Red Velvet Eksklusif',item:'ICE RED VELVET',hj:28000,hpp:9969,gm:.644,poin:500,ef:1000000,rs:.08,ket:'✅ Sangat worth it — 8% dari threshold STAY'},
  {t:'platinum',n:'Kopi Pandan Srikaya VIP',item:'KOPI PANDAN SRIKAYA',hj:28000,hpp:9096,gm:.675,poin:510,ef:1020000,rs:.09,ket:'✅ Sangat worth it — 9% dari threshold STAY'},
  {t:'platinum',n:'Mangguo VIP Treat 🥭',item:'MANGGUO',hj:22000,hpp:7337,gm:.667,poin:390,ef:780000,rs:.07,ket:'✅ Sangat worth it — 7% dari threshold STAY'},
  {t:'platinum',n:'VIP Lychee Indulgence',item:'LYCHEE TEA',hj:23000,hpp:6066,gm:.736,poin:360,ef:720000,rs:.06,ket:'✅ Sangat worth it — 6% dari threshold STAY'},
  {t:'platinum',n:'Hot Matcha Prestige',item:'ICE MATCHA',hj:28000,hpp:5718,gm:.796,poin:420,ef:840000,rs:.07,ket:'✅ Sangat worth it — 7% dari threshold STAY'},
  {t:'platinum',n:'Platinum Nasi Hongkong',item:'NASI GORENG HONGKONG',hj:30000,hpp:12229,gm:.592,poin:720,ef:1440000,rs:.12,ket:'✅ Sangat worth it — 12% dari threshold STAY'},
  {t:'platinum',n:'Beef Blackpepper — VIP Only',item:'NASI BEEF BLACKPAPPER',hj:38000,hpp:14631,gm:.615,poin:810,ef:1620000,rs:.14,ket:'✅ Sangat worth it — 14% dari threshold STAY'},
  {t:'platinum',n:'Platinum Beef Teriyaki',item:'NASI BEEF TERIYAKI',hj:38000,hpp:15223,gm:.599,poin:810,ef:1620000,rs:.14,ket:'✅ Sangat worth it — 14% dari threshold STAY'},
  {t:'platinum',n:'Platinum Mie Kaldu Double',item:'INDOMIE KALDU DOUBLE',hj:22000,hpp:8378,gm:.619,poin:600,ef:1200000,rs:.10,ket:'✅ Sangat worth it — 10% dari threshold STAY'},
  {t:'platinum',n:'Hazelnut Latte VIP',item:'ICE HAZELNUT LATTE',hj:28000,hpp:8545,gm:.695,poin:570,ef:1140000,rs:.10,ket:'✅ Sangat worth it — 10% dari threshold STAY'},
  {t:'platinum',n:'Pijar Matahari Platinum 🌟',item:'PIJAR MATAHARI',hj:28000,hpp:8667,gm:.690,poin:630,ef:1260000,rs:.10,ket:'✅ Sangat worth it — 10% dari threshold STAY'},
];
const SB_RW_DEFAULT = RW.slice();

const sbRp = v => 'Rp ' + Math.round(v).toLocaleString('id-ID');
const sbPct = v => (v*100).toFixed(1)+'%';
const sbParseRp = s => { if(!s) return 0; return parseFloat(String(s).replace(/[^0-9]/g,''))||0; };
function rpFmt(el) { let v=el.value.replace(/[^0-9]/g,''); el.value=v?'Rp '+parseInt(v).toLocaleString('id-ID'):''; }
function qcFmt(el) { const d=el.value.replace(/[^0-9]/g,''); el.value=d?'Rp '+parseInt(d).toLocaleString('id-ID'):''; }
function shortRp(v) {
  if (v >= 1000000) return 'Rp' + (v/1000000).toLocaleString('id-ID', {maximumFractionDigits:1}) + 'jt';
  if (v >= 1000)    return 'Rp' + (v/1000).toLocaleString('id-ID', {maximumFractionDigits:0}) + 'rb';
  return 'Rp' + v.toLocaleString('id-ID');
}

function updateTierDisplay() {
  const t = TIER_TH;
  const set = (id, txt) => { const el = document.getElementById(id); if(el) el.textContent = txt; };
  // Tier cards
  set('th-br-naik', '≥ ' + shortRp(t.bronze.naik) + '/thn');
  set('th-si-naik', '≥ ' + shortRp(t.silver.naik) + '/thn');
  set('th-si-stay', shortRp(t.silver.stay) + '/thn');
  set('th-go-naik', '≥ ' + shortRp(t.gold.naik) + '/thn');
  set('th-go-stay', shortRp(t.gold.stay) + '/thn');
  set('th-pl-stay', shortRp(t.platinum.stay) + '/thn');
  set('th-pl-turun', '< ' + shortRp(t.platinum.stay));
  // Alur flow
  set('flow-br-naik', '≥ ' + shortRp(t.bronze.naik) + ' ↗');
  set('flow-si-stay', 'Stay ≥ ' + shortRp(t.silver.stay));
  set('flow-si-naik', '≥ ' + shortRp(t.silver.naik) + ' ↗');
  set('flow-si-down', '↙ <' + shortRp(t.silver.stay));
  set('flow-go-stay', 'Stay ≥ ' + shortRp(t.gold.stay));
  set('flow-go-naik', '≥ ' + shortRp(t.gold.naik) + ' ↗');
  set('flow-go-down', '↙ <' + shortRp(t.gold.stay));
  set('flow-pl-stay', 'Stay ≥ ' + shortRp(t.platinum.stay));
}

function sbSpill(gm) {
  if(gm>=.45) return '<span class="sp sp-am">✅ AMAN</span>';
  if(gm>=.38) return '<span class="sp sp-ti">⚠️ MENIPIS</span>';
  return '<span class="sp sp-bo">🔴 BONCOS</span>';
}
function sbTbadge(t) {
  const m={bronze:'tb-br',silver:'tb-si',gold:'tb-go',platinum:'tb-pl'};
  const l={bronze:'🥉 Bronze',silver:'🥈 Silver',gold:'🥇 Gold',platinum:'💎 Platinum'};
  return `<span class="tbadge ${m[t]}">${l[t]}</span>`;
}
function sbGc(v) { return v>=.45?'var(--success)':v>=.38?'#f59e0b':'#ef4444'; }

function sbGoto(id, el) {
  document.querySelectorAll('.sobun-wrapper .page').forEach(p=>p.classList.remove('active'));
  document.querySelectorAll('.sobun-wrapper .nav-tab').forEach(t=>t.classList.remove('active'));
  document.getElementById('sb-page-'+id).classList.add('active');
  el.classList.add('active');
  // Simpan tab aktif agar bisa di-restore setelah refresh
  try { localStorage.setItem('taobun_active_sobun_tab', id); } catch(e) {}
  // Re-render the tab content so data is always fresh when switching
  if (id === 'master')    renderMaster();
  if (id === 'reward')    renderRw();
  if (id === 'dashboard') initDash();
  if (id === 'tiering')   calcT();
  if (id === 'profit')    calcP();
  if (id === 'promo')     cbRender();
  sbSave();
}

function sbSave() {
  const isCustomP = P.length !== SB_P.length || P.some((p,i) => !SB_P[i] || p.n !== SB_P[i].n || p.hpp !== SB_P[i].hpp);
  const isCustomRW = RW.length !== SB_RW_DEFAULT.length || RW.some((r,i) => !SB_RW_DEFAULT[i] || r.n !== SB_RW_DEFAULT[i].n || r.poin !== SB_RW_DEFAULT[i].poin);
  const d = {
    theme: document.body.getAttribute('data-theme'),
    qc_b: document.getElementById('qc-b').value,
    qc_t: document.getElementById('qc-t').value,
    t_avg: document.getElementById('t-avg').value,
    promos: promos,
    customProducts: isCustomP ? P.slice() : null,
    customRewards: isCustomRW ? RW.slice() : null,
    tierThresholds: TIER_TH,
  };
  try { localStorage.setItem(LS_KEY_SOBUN, JSON.stringify(d)); flashSave(); } catch(e) {}
}
function sbLoad() {
  try {
    const raw = localStorage.getItem(LS_KEY_SOBUN);
    if (!raw) return;
    const d = JSON.parse(raw);
    const setEl = (id, val) => { if (!val && val !== 0) return; const el = document.getElementById(id); if (el) el.value = val; };

    // Theme
    if (d.theme) {
      document.body.setAttribute('data-theme', d.theme);
      document.documentElement.setAttribute('data-theme', d.theme);
    }
    // Field bebas (tidak dikunci)
    setEl('qc-b',  d.qc_b);
    setEl('qc-t',  d.qc_t);
    setEl('t-avg', d.t_avg);
    // Promos
    if (Array.isArray(d.promos)) promos = d.promos;
    // Custom produk/reward (kalau user pernah edit manual)
    if (Array.isArray(d.customProducts) && d.customProducts.length > 0) {
      P.length = 0; d.customProducts.forEach(p => P.push(p));
    }
    if (Array.isArray(d.customRewards) && d.customRewards.length > 0) {
      RW.length = 0; d.customRewards.forEach(r => RW.push(r));
    }
    // Tier thresholds (kalau pernah diubah)
    if (d.tierThresholds) TIER_TH = d.tierThresholds;

    updateTierDisplay();
    qc(); calcT(); calcP();
  } catch(e) { console.warn('[sbLoad]', e); }
}
// ─────────────────────────────────────────────────────────
// EXCEL BASELINE LOCK — mengunci field yang datang dari Excel
// ─────────────────────────────────────────────────────────
function applyExcelBaslineLock(lock) {
  // Field-field yang dikunci ketika data berasal dari Excel
  const lockedIds = ['t-bp','t-np','p-hpp','p-bp','p-np','p-abr','p-asi','p-ago','p-apl'];
  
  lockedIds.forEach(id => {
    const el = document.getElementById(id);
    if (!el) return;
    if (lock) {
      el.setAttribute('readonly', 'readonly');
      el.setAttribute('title', 'Field ini dikunci — data berasal dari import Excel. Upload Excel baru untuk mengubah.');
      el.classList.add('xl-locked');
    } else {
      el.removeAttribute('readonly');
      el.removeAttribute('title');
      el.classList.remove('xl-locked');
    }
  });

  // Tampilkan / sembunyikan badge info Excel di Sobat Taobun
  let badge = document.getElementById('xl-baseline-badge');
  if (lock) {
    if (!badge) {
      badge = document.createElement('div');
      badge.id = 'xl-baseline-badge';
      badge.innerHTML = `
        <span style="font-size:14px">📊</span>
        <div style="flex:1">
          <div style="font-weight:800;font-size:12px">Data dari Import Excel</div>
          <div style="font-size:11px;opacity:.8">Field parameter dikunci. Upload Excel baru untuk memperbarui data.</div>
        </div>
      `;
      badge.style.cssText = `
        display:flex;align-items:center;gap:10px;
        background:linear-gradient(135deg,rgba(214,30,48,0.08),rgba(214,30,48,0.04));
        border:1.5px solid rgba(214,30,48,0.25);
        border-radius:14px;padding:12px 16px;margin-bottom:18px;
        color:var(--text);
      `;
      // Inject setelah nav, sebelum page pertama
      const firstPage = document.querySelector('.sobun-wrapper .page');
      if (firstPage) firstPage.parentNode.insertBefore(badge, firstPage);
    }
    badge.style.display = 'flex';
  } else {
    if (badge) badge.style.display = 'none';
  }
}


function flashSave() {
  const dot = document.getElementById('saveDot');
  if(!dot) return;
  dot.classList.remove('visible');
  void dot.offsetWidth;
  dot.classList.add('visible');
}

function initDash() {
  const t1 = document.getElementById('hero-total-prod');
  const t2 = document.getElementById('stat-total-prod');
  if (t1) t1.textContent = P.length;
  if (t2) t2.textContent = P.length;
  updateTierDisplay();
}

function qc() {
  const v = sbParseRp(document.getElementById('qc-b').value);
  const m = parseFloat(document.getElementById('qc-t').value)||1;
  if (!v) {
    document.getElementById('qc-res').style.display='none';
    document.getElementById('qc-ph').style.display='block';
    return;
  }
  document.getElementById('qc-res').style.display='block';
  document.getElementById('qc-ph').style.display='none';
  const base=Math.floor(v/2000), tot=Math.floor(base*m), val=tot*20;
  document.getElementById('qc-base').textContent=base+' poin';
  document.getElementById('qc-mult').textContent=m+'×';
  document.getElementById('qc-tot').textContent=tot+' poin';
  document.getElementById('qc-val').textContent=sbRp(val);
  document.getElementById('qc-col').textContent=sbPct(v>0?val/v:0);
}

let mpF = 'semua';
function mpFilter(el) {
  mpF = el.dataset.f;
  document.querySelectorAll('#mp-fpr .fp').forEach(x=>x.classList.remove('active'));
  el.classList.add('active');
  renderMaster();
}
function renderMaster() {
  const q = document.getElementById('mp-q').value.toLowerCase();
  const data = P.filter(p=>{
    if (q&&!p.n.toLowerCase().includes(q)) return false;
    if (mpF==='Pareto A') return p.par==='Pareto A';
    if (mpF==='Pareto B') return p.par==='Pareto B';
    if (mpF==='Non-Pareto') return p.par==='Non-Pareto';
    if (mpF==='bisa') return p.red==='BISA';
    if (mpF==='tidak') return p.red==='TIDAK';
    return true;
  });
  const tc={'Pareto A':'tag-a','Pareto B':'tag-b','Non-Pareto':'tag-np'};
  document.getElementById('master-tb').innerHTML = data.map((p,i)=>`
    <tr>
      <td class="mono" style="color:var(--muted)">${i+1}</td>
      <td style="font-weight:700;white-space:nowrap">${p.n}</td>
      <td class="tr mono">${sbRp(p.hpp)}</td>
      <td class="tr mono">${sbRp(p.hj)}</td>
      <td class="tr mono">${sbRp(p.gm)}</td>
      <td class="tr mono" style="color:${sbGc(p.gmp)}">${sbPct(p.gmp)}</td>
      <td class="tc2"><span class="tag ${tc[p.par]||'tag-np'}">${p.par}</span></td>
      <td class="tr mono">${p.poin??'—'}</td>
      <td>${p.red==='BISA'?'<span class="sp sp-am">✅ Bisa</span>':'<span class="sp sp-bo">❌ Tidak</span>'}</td>
    </tr>`).join('');
  document.getElementById('mp-count').textContent=`${data.length} dari ${P.length} produk ditampilkan`;
}

function calcT() {
  const BP=sbParseRp(document.getElementById('t-bp').value)||2000;
  const NP=sbParseRp(document.getElementById('t-np').value)||20;
  const avg=sbParseRp(document.getElementById('t-avg').value)||0;
  document.getElementById('t-sim-tb').innerHTML = SB_TIERS.map(t=>{
    const poin=avg>0?Math.floor(avg/BP*t.mult):0, val=poin*NP, col=avg>0?val/avg:0;
    const cc=col<=.02?'var(--success)':col<=.03?'#f59e0b':'#ef4444';
    return `<tr>
      <td><span class="tbadge ${t.css}">${t.lbl}</span></td>
      <td class="tr mono">${avg>0?poin+' poin':'—'}</td>
      <td class="tr mono">${avg>0?sbRp(val):'—'}</td>
      <td class="tr mono" style="color:${avg>0?cc:'var(--muted)'}">${avg>0?sbPct(col):'—'}</td>
    </tr>`;
  }).join('');
  SB_TIERS.forEach(t=>{
    const poin=avg>0?Math.floor(avg/sbParseRp(document.getElementById('t-bp').value||'2000')*t.mult):0;
    const el=document.getElementById('tci-'+t.k);
    if(el) el.innerHTML=`
      <div class="tc-row"><span>Poin per Transaksi</span><span>${avg>0?poin+' poin':'—'}</span></div>
      <div class="tc-row"><span>Nilai Poin</span><span>${avg>0?sbRp(poin*NP):'—'}</span></div>
      <div class="tc-row"><span>Cost of Loyalty</span><span>${avg>0?sbPct(avg>0?poin*NP/avg:0):'—'}</span></div>`;
  });
}

function calcP() {
  const hp=(parseFloat(document.getElementById('p-hpp').value)||36.55)/100;
  const BP=sbParseRp(document.getElementById('p-bp').value)||2000;
  const NP=sbParseRp(document.getElementById('p-np').value)||20;
  const avgs={
    bronze:sbParseRp(document.getElementById('p-abr').value),
    silver:sbParseRp(document.getElementById('p-asi').value),
    gold:sbParseRp(document.getElementById('p-ago').value),
    platinum:sbParseRp(document.getElementById('p-apl').value),
  };
  document.getElementById('profit-tb').innerHTML = SB_TIERS.map(t=>{
    const avg=avgs[t.k]||0;
    const eh=avg*hp, gma=avg-eh, poin=Math.floor(avg/BP)*t.mult, bp2=poin*NP, gmf=gma-bp2, gmp=avg>0?gmf/avg:0;
    return `<tr>
      <td><span class="tbadge ${t.css}">${t.lbl}</span></td>
      <td class="tr mono">${avg>0?sbRp(avg):'—'}</td>
      <td class="tr mono">${avg>0?sbRp(eh):'—'}</td>
      <td class="tr mono">${avg>0?sbRp(gma):'—'}</td>
      <td class="tr mono" style="color:#ef4444">${avg>0?sbRp(bp2):'—'}</td>
      <td class="tr mono">${avg>0?sbRp(gmf):'—'}</td>
      <td class="tr mono" style="color:${sbGc(gmp)}">${avg>0?sbPct(gmp):'—'}</td>
      <td>${avg>0?sbSpill(gmp):'—'}</td>
    </tr>`;
  }).join('');
}

let rwF = 'semua';
function rwFilter(el) {
  rwF = el.dataset.f;
  document.querySelectorAll('#rw-fpr .fp').forEach(x=>x.classList.remove('active'));
  el.classList.add('active');
  renderRw();
}
function renderRw() {
  const data = rwF==='semua'?RW:RW.filter(r=>r.t===rwF);
  document.getElementById('rw-tb').innerHTML = data.map(r=>{
    const rc=r.rs!==null?(r.rs<=.15?'var(--success)':r.rs<=.30?'#f59e0b':'#ef4444'):'var(--muted)';
    return `<tr>
      <td>${sbTbadge(r.t)}</td>
      <td style="font-weight:700;white-space:nowrap">${r.n}</td>
      <td style="white-space:nowrap">${r.item}</td>
      <td class="tr mono">${sbRp(r.hj)}</td>
      <td class="tr mono">${sbRp(r.hpp)}</td>
      <td class="tr mono" style="color:${sbGc(r.gm)}">${sbPct(r.gm)}</td>
      <td class="tr mono" style="color:var(--primary)">${r.poin}</td>
      <td class="tr mono">${sbRp(r.ef)}</td>
      <td class="tr mono" style="color:${rc}">${r.rs!==null?sbPct(r.rs):'—'}</td>
      <td style="font-size:11px;color:var(--text);min-width:160px">${r.ket}</td>
    </tr>`;
  }).join('');
}

// Combobox
let cbFocus = -1, selP = null;
function cbRender(q='') {
  const dd=document.getElementById('cb-dd');
  // Show/hide no-products warning
  const noBanner = document.getElementById('promo-no-products-banner');
  const cwEl = document.getElementById('cw');
  if (P.length === 0) {
    if (noBanner) noBanner.classList.add('show');
    if (cwEl) cwEl.style.display = 'none';
    return;
  } else {
    if (noBanner) noBanner.classList.remove('show');
    if (cwEl) cwEl.style.display = '';
  }
  const items=P.filter(p=>!q||p.n.toLowerCase().includes(q.toLowerCase()));
  if(!items.length){dd.innerHTML='<div class="co-empty">Produk tidak ditemukan</div>';return;}
  dd.innerHTML=items.map(p=>`
    <div class="co" data-n="${p.n}" onclick="cbSel('${p.n.replace(/'/g,"\\'")}')">
      ${p.n}<div class="co-sub">${sbRp(p.hj)} · HPP ${sbRp(p.hpp)} · GM ${sbPct(p.gmp)}</div>
    </div>`).join('');
  cbFocus=-1;
}
function cbOpen(){document.getElementById('cb-dd').classList.add('open');cbRender(document.getElementById('pr-inp').value);}
function cbFilter(){document.getElementById('cb-dd').classList.add('open');cbRender(document.getElementById('pr-inp').value);}
function cbSel(name){
  document.getElementById('pr-inp').value=name;
  document.getElementById('pr-val').value=name;
  document.getElementById('cb-dd').classList.remove('open');
  selP=P.find(p=>p.n===name);
  if(selP){
    document.getElementById('pr-hj').value=sbRp(selP.hj);
    document.getElementById('pr-hpp').value=sbRp(selP.hpp);
    const poinWrap=document.getElementById('pr-poin-wrap');
    const poinEl=document.getElementById('pr-poin');
    if(selP.poin!=null){poinEl.value=selP.poin.toLocaleString('id-ID')+' Poin';poinWrap.style.display='';}
    else{poinWrap.style.display='none';}
  }
  preLive();
}
function cbKey(e){
  const dd=document.getElementById('cb-dd'), opts=dd.querySelectorAll('.co');
  if(e.key==='ArrowDown'){e.preventDefault();cbFocus=Math.min(cbFocus+1,opts.length-1);opts.forEach((o,i)=>o.classList.toggle('foc',i===cbFocus));if(opts[cbFocus])opts[cbFocus].scrollIntoView({block:'nearest'});}
  else if(e.key==='ArrowUp'){e.preventDefault();cbFocus=Math.max(cbFocus-1,0);opts.forEach((o,i)=>o.classList.toggle('foc',i===cbFocus));}
  else if(e.key==='Enter'){e.preventDefault();if(cbFocus>=0&&opts[cbFocus])opts[cbFocus].click();else dd.classList.remove('open');}
  else if(e.key==='Escape')dd.classList.remove('open');
}
document.addEventListener('click',e=>{
  const cw=document.getElementById('cw');
  if(cw&&!cw.contains(e.target))document.getElementById('cb-dd').classList.remove('open');
});

// Promo simulasi
let promos = [];
function prTipe(){
  const t=document.getElementById('pr-tipe').value, dw=document.getElementById('pr-dw'), dl=document.getElementById('pr-dl');
  const disEl=document.getElementById('pr-dis');
  if(t==='gratis'){dw.style.display='none';}
  else{
    dw.style.display='flex';
    if(t==='drp'){dl.textContent='Nominal Diskon (Rp)';disEl.placeholder='Rp 0';}
    else{dl.textContent='Diskon (%)';disEl.placeholder='contoh: 10';}
    disEl.value='';
  }
}
function prDisInput(el){
  const t=document.getElementById('pr-tipe').value;
  if(t==='drp') rpFmt(el);
}
function calcV(prod,qty,tipe,disRaw,mt){
  if(!prod)return null;
  const hpp=prod.hpp*qty, rev=mt;
  let dis=0;
  if(tipe==='drp')dis=sbParseRp(disRaw)||0;
  else if(tipe==='dpct')dis=prod.hj*qty*((parseFloat(disRaw)||0)/100);
  const net=rev-hpp-dis, gm=rev>0?net/rev:0;
  return{hpp,rev,net,gm};
}
function preLive(){
  if(!selP){document.getElementById('pr-prev').style.display='none';return;}
  const qty=parseInt(document.getElementById('pr-qty').value)||1;
  const tipe=document.getElementById('pr-tipe').value;
  const dis=document.getElementById('pr-dis').value;
  const mt=sbParseRp(document.getElementById('pr-mt').value)||0;
  const v=calcV(selP,qty,tipe,dis,mt);
  if(!v)return;
  document.getElementById('pr-prev').style.display='block';
  document.getElementById('pv-hpp').textContent=sbRp(v.hpp);
  document.getElementById('pv-rev').textContent=sbRp(v.rev);
  document.getElementById('pv-net').textContent=sbRp(v.net);
  const gmEl=document.getElementById('pv-gm');
  gmEl.textContent=sbPct(v.gm);gmEl.style.color=sbGc(v.gm);
  document.getElementById('pv-st').innerHTML=sbSpill(v.gm);
}
function addPromo(){
  if(!selP){
    xlToast('⚠️ Pilih produk terlebih dahulu dari dropdown.');
    document.getElementById('pr-inp').classList.add('input-error');
    setTimeout(() => document.getElementById('pr-inp').classList.remove('input-error'), 2000);
    return;
  }
  const nm=document.getElementById('pr-nm').value.trim()||'Promo #'+(promos.length+1);
  const qty=parseInt(document.getElementById('pr-qty').value)||1;
  const tipe=document.getElementById('pr-tipe').value;
  const dis=document.getElementById('pr-dis').value;
  const mt=sbParseRp(document.getElementById('pr-mt').value)||0;
  const v=calcV(selP,qty,tipe,dis,mt);
  if(!v)return;
  const tl={gratis:'🎁 Gratis',drp:'💸 Diskon Rp',dpct:'📉 Diskon %'};
  promos.push({nm,prod:selP.n,qty,tipeRaw:tipe,tipe:tl[tipe],disRaw:dis,poin:selP.poin,...v});
  renderPromo();
  document.getElementById('pr-nm').value='';
  xlToast('✅ Promo "'+nm+'" berhasil ditambahkan.');
  sbSave();
}
function delPromo(i){promos.splice(i,1);renderPromo();sbSave();}
function renderPromo(){
  document.getElementById('pr-empty').style.display=promos.length?'none':'block';
  document.getElementById('pr-tw').style.display=promos.length?'block':'none';
  document.getElementById('pr-tb').innerHTML=promos.map((pr,i)=>`
    <tr>
      <td class="mono" style="color:var(--muted)">${i+1}</td>
      <td style="font-weight:700;white-space:nowrap">${pr.nm}</td>
      <td style="white-space:nowrap">${pr.prod}</td>
      <td class="tc2 mono">${pr.qty}</td>
      <td style="font-size:12px;white-space:nowrap">${pr.tipe}${pr.tipeRaw==='drp'&&pr.disRaw?' <span class="mono" style="font-size:11px;color:var(--muted)">(-'+sbRp(sbParseRp(pr.disRaw))+')</span>':pr.tipeRaw==='dpct'&&pr.disRaw?' <span class="mono" style="font-size:11px;color:var(--muted)">(-'+parseFloat(pr.disRaw)+'%)</span>':''}</td>
      <td class="tr mono" style="font-weight:800;color:${pr.tipeRaw==='gratis'?'var(--primary)':pr.tipeRaw==='drp'?'var(--success)':'#f59e0b'}">${(()=>{if(pr.tipeRaw==='gratis')return pr.poin!=null?pr.poin.toLocaleString('id-ID')+' Poin':'—';if(pr.tipeRaw==='drp')return sbRp(sbParseRp(pr.disRaw));if(pr.tipeRaw==='dpct')return (parseFloat(pr.disRaw)||0)+'%';return '—';})()}</td>
      <td class="tr mono">${sbRp(pr.hpp)}</td>
      <td class="tr mono">${sbRp(pr.rev)}</td>
      <td class="tr mono" style="color:${pr.net>=0?'var(--success)':'#ef4444'}">${sbRp(pr.net)}</td>
      <td class="tr mono" style="color:${sbGc(pr.gm)}">${sbPct(pr.gm)}</td>
      <td>${sbSpill(pr.gm)}</td>
      <td><button class="btn-del" onclick="delPromo(${i})">✕</button></td>
    </tr>`).join('');
}
// Import Excel (Sobun)
let xlParsed = null;
function xlOpen(){document.getElementById('xl-overlay').classList.add('open');xlReset();}
function xlClose(){document.getElementById('xl-overlay').classList.remove('open');xlParsed=null;}
function xlOverlayClick(e){if(e.target===document.getElementById('xl-overlay'))xlClose();}
function xlReset(){
  xlParsed=null;
  document.getElementById('xl-drop-txt').textContent='Klik atau drag & drop file Excel di sini';
  document.getElementById('xl-drop').classList.remove('drag');
  document.getElementById('xl-progress').classList.remove('show');
  document.getElementById('xl-err').classList.remove('show');
  document.getElementById('xl-preview').classList.remove('show');
  document.getElementById('xl-apply-btn').disabled=true;
  document.getElementById('xl-prog-fill').style.width='0%';
}
function xlDrag(e,over){e.preventDefault();document.getElementById('xl-drop').classList.toggle('drag',over);}
function xlDropFile(e){e.preventDefault();document.getElementById('xl-drop').classList.remove('drag');const f=e.dataTransfer.files[0];if(f)xlProcess(f);}
function xlFileInput(e){const f=e.target.files[0];if(f)xlProcess(f);}
function xlProgress(pct,msg){document.getElementById('xl-progress').classList.add('show');document.getElementById('xl-prog-fill').style.width=pct+'%';document.getElementById('xl-prog-txt').textContent=msg;}
function xlShowErr(msg){const el=document.getElementById('xl-err');el.textContent='⚠️ '+msg;el.classList.add('show');}
function xlProcess(file){
  loadXlsx(function() { _doXlProcess(file); });
}
function _doXlProcess(file){
  if(!file.name.match(/\.xlsx?$/i)){xlShowErr('File harus berformat .xlsx atau .xls');return;}
  document.getElementById('xl-drop-txt').textContent='📂 '+file.name;
  document.getElementById('xl-err').classList.remove('show');
  document.getElementById('xl-preview').classList.remove('show');
  document.getElementById('xl-apply-btn').disabled=true;
  xlProgress(10,'Membaca file...');
  const reader=new FileReader();
  reader.onload=function(e){
    try{
      xlProgress(40,'Parsing Excel...');
      const data=new Uint8Array(e.target.result), wb=XLSX.read(data,{type:'array'});
      xlProgress(70,'Mengolah data...');
      xlParsed=xlExtract(wb);
      xlProgress(100,'Selesai!');
      xlShowPreview(xlParsed);
      document.getElementById('xl-apply-btn').disabled=false;
    }catch(err){xlShowErr('Gagal membaca file: '+err.message);}
  };
  reader.readAsArrayBuffer(file);
}
function xlExtract(wb){
  const result={};

  // ── Sheet 1: Master Produk ──
  if(wb.SheetNames.includes('Master Produk')){
    const ws=wb.Sheets['Master Produk'],rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:''}),products=[];
    for(let i=1;i<rows.length;i++){
      const r=rows[i];
      const nama=String(r[0]||'').trim();
      if(!nama)continue;
      const hpp=parseFloat(r[1])||0;
      const hj=parseFloat(r[2])||0;
      const gm=hj-hpp;
      const gmp=hj>0?gm/hj:0;
      const par=String(r[5]||'Non-Pareto').trim();
      const redRaw=String(r[6]||'').toUpperCase();
      const red=redRaw.includes('BISA')?'BISA':(redRaw.includes('TIDAK')?'TIDAK':((gmp>=0.6&&hj<=30000)?'BISA':'TIDAK'));
      const poinRaw=parseFloat(r[7]);
      const mtRaw=parseFloat(r[8]);
      const poin=red==='BISA'?(isNaN(poinRaw)?Math.ceil(hpp/20/5)*5:poinRaw):null;
      const mt=red==='BISA'?(isNaN(mtRaw)?Math.ceil((poin*20+hpp)/(1-0.3655)/1000)*1000:mtRaw):null;
      products.push({n:nama,hpp,hj,gm,gmp,par,red,poin,mt});
    }
    result.products=products;
  }

  // ── Sheet 2: Skema Tiering ──
  if(wb.SheetNames.includes('Skema Tiering (Max 2%)')){
    const ws=wb.Sheets['Skema Tiering (Max 2%)'],rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
    result.tiering={bp:parseFloat(rows[1]?.[1])||2000,np:parseFloat(rows[2]?.[1])||20};
  }

  // ── Sheet 3: Simulasi Profitabilitas ──
  if(wb.SheetNames.includes('Simulasi Profitabilitas')){
    const ws=wb.Sheets['Simulasi Profitabilitas'],rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
    const hppPct=parseFloat(rows[1]?.[1]);
    if(!isNaN(hppPct)) result.hppPct=hppPct<=1?hppPct*100:hppPct;
    const avgBronze=parseFloat(rows[5]?.[1]);
    const avgSilver=parseFloat(rows[6]?.[1]);
    const avgGold=parseFloat(rows[7]?.[1]);
    const avgPlatinum=parseFloat(rows[8]?.[1]);
    result.avgStruk={
      bronze:  isNaN(avgBronze)?0:avgBronze,
      silver:  isNaN(avgSilver)?0:avgSilver,
      gold:    isNaN(avgGold)?0:avgGold,
      platinum:isNaN(avgPlatinum)?0:avgPlatinum,
    };
    // Tier thresholds: rows 13-16 (index), cols 2=NAIK, 3=STAY
    const tierRows = [rows[13],rows[14],rows[15],rows[16]];
    const tierKeys = ['bronze','silver','gold','platinum'];
    const thresholds = {};
    tierRows.forEach((r,i)=>{
      if(!r) return;
      const naik=parseFloat(r[2])||0;
      const stay=parseFloat(r[3])||0;
      if(naik||stay) thresholds[tierKeys[i]]={naik,stay};
    });
    if(Object.keys(thresholds).length) result.thresholds=thresholds;
  }

  // ── Sheet 4: Simulasi Misi & Promo ──
  if(wb.SheetNames.includes('Simulasi Misi & Promo')){
    const ws=wb.Sheets['Simulasi Misi & Promo'],rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
    const promoData=[];
    for(let i=4;i<rows.length;i++){
      const r=rows[i];
      const no=parseFloat(r[0]);
      const nama=String(r[1]||'').trim();
      if(!nama||isNaN(no))continue;
      const prod=String(r[2]||'').trim();
      const qty=parseFloat(r[3])||1;
      const tipeRaw=String(r[4]||'').trim().toLowerCase();
      const tipe=tipeRaw.includes('gratis')?'gratis':tipeRaw.includes('diskon rp')?'drp':'dpct';
      const hj=parseFloat(r[5])||0;
      const hpp=parseFloat(r[6])||0;
      const mt=parseFloat(r[7])||0;
      const disNominal=parseFloat(r[8])||0;
      const poin=String(r[9]||'').replace(/[^0-9]/g,'')||null;
      const rev=parseFloat(r[12])||mt;
      const net=parseFloat(r[13])||0;
      const gm=parseFloat(r[14])||0;
      promoData.push({
        nm:nama, prod, qty,
        tipeRaw:tipe,
        tipe:tipe==='gratis'?'🎁 Gratis':tipe==='drp'?'💸 Diskon Rp':'📉 Diskon %',
        disRaw:tipe==='drp'?('Rp '+disNominal.toLocaleString('id-ID')):(tipe==='dpct'?String(disNominal):''),
        poin:poin?parseInt(poin):null,
        hpp:hpp*qty, rev, net, gm
      });
    }
    result.promos=promoData;
  }

  // ── Sheet 5: Simulasi Reward ──
  if(wb.SheetNames.includes('Simulasi Reward')){
    const ws=wb.Sheets['Simulasi Reward'],rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
    const rwData=[];
    // Data rows start at index 14 (after header/legend rows)
    // Cols: 0=Nama Promo, 1=Item Reward, 2=Harga Jual, 3=HPP, 4=GM%, 5=Poin, 6=Belanja Efektif, 7=Rasio STAY, 8=Nilai%, 9=~Tx, 10=Keterangan
    // Tier headers are at rows where col1 is empty and col0 contains tier emoji
    let currentTier = 'bronze';
    const tierMap = {'🥉':'bronze','🥈':'silver','🥇':'gold','💎':'platinum'};
    for(let i=13;i<rows.length;i++){
      const r=rows[i];
      const col0=String(r[0]||'').trim();
      // Detect tier header rows
      const tierMatch=Object.keys(tierMap).find(k=>col0.startsWith(k));
      if(tierMatch){currentTier=tierMap[tierMatch];continue;}
      const nama=col0;
      if(!nama||!r[1])continue;
      const item=String(r[1]||'').trim();
      if(!item)continue;
      const hj=parseFloat(r[2])||0;
      const hpp=parseFloat(r[3])||0;
      const gm=parseFloat(r[4])||0;
      const poin=parseFloat(r[5])||0;
      const ef=parseFloat(r[6])||0;
      const rs=parseFloat(r[7])||null;
      const ket=String(r[10]||'').trim();
      if(hj&&hpp&&poin) rwData.push({t:currentTier,n:nama,item,hj,hpp,gm,poin,ef,rs:isNaN(rs)?null:rs,ket});
    }
    if(rwData.length) result.rewards=rwData;
  }

  return result;
}
function xlShowPreview(d){
  const rp = v => 'Rp '+Math.round(v).toLocaleString('id-ID');
  const rows=[];

  // ── Sheet 1: Master Produk ──
  if(d.products && d.products.length){
    const bisaCount  = d.products.filter(p=>p.red==='BISA').length;
    const tidakCount = d.products.filter(p=>p.red==='TIDAK').length;
    const paretoA    = d.products.filter(p=>p.par==='Pareto A').length;
    rows.push({label:'📦 Master Produk',   val: d.products.length+' produk', section:true});
    rows.push({label:'  ✅ Bisa Redeem',    val: bisaCount+' produk'});
    rows.push({label:'  ❌ Tidak Redeem',   val: tidakCount+' produk'});
    rows.push({label:'  🏆 Pareto A',       val: paretoA+' produk'});
  }

  // ── Sheet 2: Skema Tiering ──
  if(d.tiering){
    rows.push({label:'⚙️ Skema Tiering', val:'', section:true});
    rows.push({label:'  Syarat Belanja → 1 Poin', val: rp(d.tiering.bp)});
    rows.push({label:'  Nilai Tukar 1 Poin',       val: rp(d.tiering.np)});
  }

  // ── Sheet 3: Simulasi Profitabilitas ──
  if(d.hppPct !== undefined){
    rows.push({label:'📊 Simulasi Profitabilitas', val:'', section:true});
    rows.push({label:'  Asumsi HPP (%)', val: d.hppPct.toFixed(2)+'%'});
  }
  if(d.avgStruk){
    const a=d.avgStruk;
    if(a.bronze)   rows.push({label:'  Avg Struk 🥉 Bronze',   val: rp(a.bronze)});
    if(a.silver)   rows.push({label:'  Avg Struk 🥈 Silver',   val: rp(a.silver)});
    if(a.gold)     rows.push({label:'  Avg Struk 🥇 Gold',     val: rp(a.gold)});
    if(a.platinum) rows.push({label:'  Avg Struk 💎 Platinum', val: rp(a.platinum)});
  }
  if(d.thresholds){
    const t=d.thresholds;
    if(t.silver)   rows.push({label:'  Silver: Naik/Stay',  val: rp(t.silver.naik)+' / '+rp(t.silver.stay)});
    if(t.gold)     rows.push({label:'  Gold: Naik/Stay',    val: rp(t.gold.naik)+' / '+rp(t.gold.stay)});
    if(t.platinum) rows.push({label:'  Platinum: Naik/Stay',val: rp(t.platinum.naik)+' / '+rp(t.platinum.stay)});
  }

  // ── Sheet 4: Simulasi Misi & Promo ──
  if(d.promos && d.promos.length){
    const gratis = d.promos.filter(p=>p.tipeRaw==='gratis').length;
    const diskon = d.promos.filter(p=>p.tipeRaw!=='gratis').length;
    const aman   = d.promos.filter(p=>p.gm>=0.45).length;
    const boncos = d.promos.filter(p=>p.gm<0.38).length;
    rows.push({label:'🎯 Simulasi Misi & Promo', val: d.promos.length+' promo', section:true});
    rows.push({label:'  🎁 Gratis', val: gratis+' promo'});
    rows.push({label:'  💸 Diskon', val: diskon+' promo'});
    rows.push({label:'  ✅ Aman (GM ≥45%)', val: aman+' promo'});
    if(boncos) rows.push({label:'  🔴 Boncos (GM <38%)', val: boncos+' promo', warn:true});
  }

  // ── Sheet 5: Simulasi Reward ──
  if(d.rewards && d.rewards.length){
    const byTier = {bronze:0,silver:0,gold:0,platinum:0};
    d.rewards.forEach(r=>{ if(byTier[r.t]!==undefined) byTier[r.t]++; });
    rows.push({label:'🏆 Simulasi Reward', val: d.rewards.length+' reward', section:true});
    if(byTier.bronze)   rows.push({label:'  🥉 Bronze',   val: byTier.bronze+' reward'});
    if(byTier.silver)   rows.push({label:'  🥈 Silver',   val: byTier.silver+' reward'});
    if(byTier.gold)     rows.push({label:'  🥇 Gold',     val: byTier.gold+' reward'});
    if(byTier.platinum) rows.push({label:'  💎 Platinum', val: byTier.platinum+' reward'});
  }

  const el=document.getElementById('xl-preview');
  document.getElementById('xl-preview-rows').innerHTML=rows.map(r=>`
    <div class="xl-prow${r.section?' xl-prow-section':''}${r.warn?' xl-prow-warn':''}">
      <span>${r.label}</span>
      <span style="font-weight:${r.section?'800':'700'};color:${r.warn?'var(--danger)':r.section?'var(--primary)':'var(--text)'}">${r.val}</span>
    </div>`).join('');
  el.classList.add('show');
}
function xlApply(){
  if(!xlParsed)return;
  // Hide import modal WITHOUT nulling xlParsed — password gate needs it
  document.getElementById('xl-overlay').classList.remove('open');
  // Show password confirmation modal
  document.getElementById('pwImportOverlay').classList.add('open');
  const inp = document.getElementById('pwImportInput');
  inp.value = '';
  inp.type = 'password';
  document.getElementById('pwImportToggleBtn').textContent = '👁';
  document.getElementById('pwImportErrMsg').textContent = '';
  inp.classList.remove('pw-error');
  setTimeout(()=>inp.focus(),150);
}
function xlToast(msg){const el=document.getElementById('xl-toast');el.textContent=msg;el.classList.add('show');setTimeout(()=>el.classList.remove('show'),3500);}

// ═══════════════════════════════════════════
// BRANCH COMPARISON
// ═══════════════════════════════════════════
let bcChart = null;

function bcHandleInput(el) {
  el.value = formatNumber(el.value);
  bcCalculate();
}
function bcRestoreIfEmpty(el) {
  if (el.value === '') { el.value = '0'; bcCalculate(); }
}

function bcGet(id) { return cleanNumber(document.getElementById(id)?.value || '0'); }

function bcRp(n) { return 'Rp ' + Math.round(n).toLocaleString('id-ID'); }
function bcPct(n) { return (n * 100).toFixed(1) + '%'; }

function bcWinnerBadge(winner) {
  if (winner === 'tie') return '<span class="branch-winner winner-tie">= Sama</span>';
  if (winner === 'raya') return '<span class="branch-winner winner-raya">🟣 Raya</span>';
  if (winner === 'perdos') return '<span class="branch-winner winner-perdos">🟢 Perdos</span>';
  return '<span class="branch-winner winner-kemboja">🟡 Kemboja</span>';
}

function bcDelta(a, b, fmt = 'rp') {
  const diff = a - b;
  if (diff === 0) return '<span class="delta-neu">—</span>';
  const abs = Math.abs(diff);
  const label = fmt === 'pct' ? bcPct(abs)
              : fmt === 'num' ? abs.toLocaleString('id-ID')
              : bcRp(abs);
  return diff > 0
    ? `<span class="delta-pos">+${label}</span>`
    : `<span class="delta-neg">-${label}</span>`;
}

function bcBestOf3(r, k, p, higherIsBetter = true) {
  const vals = [['raya', r], ['kemboja', k], ['perdos', p]];
  if (higherIsBetter) vals.sort((a,b) => b[1] - a[1]);
  else vals.sort((a,b) => a[1] - b[1]);
  if (vals[0][1] === vals[1][1]) return 'tie';
  return vals[0][0];
}

function bcCalculate() {
  const r = {
    rev:    bcGet('br-raya-rev'),
    cust:   bcGet('br-raya-cust'),
    tx:     bcGet('br-raya-tx'),
    basket: bcGet('br-raya-basket'),
    new_c:  bcGet('br-raya-new'),
    cogs:   bcGet('br-raya-cogs'),
  };
  const k = {
    rev:    bcGet('br-kemb-rev'),
    cust:   bcGet('br-kemb-cust'),
    tx:     bcGet('br-kemb-tx'),
    basket: bcGet('br-kemb-basket'),
    new_c:  bcGet('br-kemb-new'),
    cogs:   bcGet('br-kemb-cogs'),
  };
  const p = {
    rev:    bcGet('br-perd-rev'),
    cust:   bcGet('br-perd-cust'),
    tx:     bcGet('br-perd-tx'),
    basket: bcGet('br-perd-basket'),
    new_c:  bcGet('br-perd-new'),
    cogs:   bcGet('br-perd-cogs'),
  };

  // Derived metrics
  r.gm = r.rev > 0 ? (r.rev - r.cogs) / r.rev : 0;
  k.gm = k.rev > 0 ? (k.rev - k.cogs) / k.rev : 0;
  p.gm = p.rev > 0 ? (p.rev - p.cogs) / p.rev : 0;
  r.atv = r.tx > 0 ? r.rev / r.tx : 0;
  k.atv = k.tx > 0 ? k.rev / k.tx : 0;
  p.atv = p.tx > 0 ? p.rev / p.tx : 0;
  r.retRate = r.cust > 0 ? (r.cust - r.new_c) / r.cust : 0;
  k.retRate = k.cust > 0 ? (k.cust - k.new_c) / k.cust : 0;
  p.retRate = p.cust > 0 ? (p.cust - p.new_c) / p.cust : 0;
  r.revPerCust = r.cust > 0 ? r.rev / r.cust : 0;
  k.revPerCust = k.cust > 0 ? k.rev / k.cust : 0;
  p.revPerCust = p.cust > 0 ? p.rev / p.cust : 0;

  // KPI cards
  const totalRev = r.rev + k.rev + p.rev;
  document.getElementById('bc-total-rev').textContent = bcRp(totalRev);
  document.getElementById('bc-total-cust').textContent = (r.cust + k.cust + p.cust).toLocaleString('id-ID');

  const revs = [['Raya','#6366f1',r.rev], ['Kemboja','#f59e0b',k.rev], ['Perdos','#10b981',p.rev]];
  revs.sort((a,b) => b[2] - a[2]);
  document.getElementById('bc-rev-leader').innerHTML = `<span style="color:${revs[0][1]}">${revs[0][0]}</span>`;
  document.getElementById('bc-rev-diff').textContent = revs[0][2] > revs[1][2] ? `Unggul ${bcRp(revs[0][2] - revs[1][2])} vs no.2` : 'Semua sama';

  const gms = [['Raya','#6366f1',r.gm], ['Kemboja','#f59e0b',k.gm], ['Perdos','#10b981',p.gm]];
  gms.sort((a,b) => b[2] - a[2]);
  document.getElementById('bc-gm-leader').innerHTML = `<span style="color:${gms[0][1]}">${gms[0][0]} ${bcPct(gms[0][2])}</span>`;
  document.getElementById('bc-gm-diff').textContent = `Selisih vs no.2: ${bcPct(Math.abs(gms[0][2] - gms[1][2]))}`;

  // Comparison table
  const metrics = [
    { label: 'Revenue',           rVal: bcRp(r.rev),           kVal: bcRp(k.rev),           pVal: bcRp(p.rev),           winner: bcBestOf3(r.rev, k.rev, p.rev) },
    { label: 'Gross Profit',      rVal: bcRp(r.rev-r.cogs),    kVal: bcRp(k.rev-k.cogs),    pVal: bcRp(p.rev-p.cogs),    winner: bcBestOf3(r.rev-r.cogs, k.rev-k.cogs, p.rev-p.cogs) },
    { label: 'Gross Margin %',    rVal: bcPct(r.gm),           kVal: bcPct(k.gm),           pVal: bcPct(p.gm),           winner: bcBestOf3(r.gm, k.gm, p.gm) },
    { label: 'Active Customers',  rVal: r.cust.toLocaleString('id-ID'), kVal: k.cust.toLocaleString('id-ID'), pVal: p.cust.toLocaleString('id-ID'), winner: bcBestOf3(r.cust, k.cust, p.cust) },
    { label: 'Transactions',      rVal: r.tx.toLocaleString('id-ID'),   kVal: k.tx.toLocaleString('id-ID'),   pVal: p.tx.toLocaleString('id-ID'),   winner: bcBestOf3(r.tx, k.tx, p.tx) },
    { label: 'Avg Transaction',   rVal: bcRp(r.atv),           kVal: bcRp(k.atv),           pVal: bcRp(p.atv),           winner: bcBestOf3(r.atv, k.atv, p.atv) },
    { label: 'Avg Basket',        rVal: bcRp(r.basket),        kVal: bcRp(k.basket),        pVal: bcRp(p.basket),        winner: bcBestOf3(r.basket, k.basket, p.basket) },
    { label: 'Rev / Customer',    rVal: bcRp(r.revPerCust),    kVal: bcRp(k.revPerCust),    pVal: bcRp(p.revPerCust),    winner: bcBestOf3(r.revPerCust, k.revPerCust, p.revPerCust) },
    { label: 'New Customers',     rVal: r.new_c.toLocaleString('id-ID'), kVal: k.new_c.toLocaleString('id-ID'), pVal: p.new_c.toLocaleString('id-ID'), winner: bcBestOf3(r.new_c, k.new_c, p.new_c) },
    { label: 'Retention Rate',    rVal: bcPct(r.retRate),      kVal: bcPct(k.retRate),      pVal: bcPct(p.retRate),      winner: bcBestOf3(r.retRate, k.retRate, p.retRate) },
    { label: 'COGS / HPP',        rVal: bcRp(r.cogs),          kVal: bcRp(k.cogs),          pVal: bcRp(p.cogs),          winner: bcBestOf3(r.cogs, k.cogs, p.cogs, false) },
  ];

  document.getElementById('bc-table-body').innerHTML = metrics.map(m => `
    <tr>
      <td class="bc-metric-name">${m.label}</td>
      <td class="bc-val-raya">${m.rVal}</td>
      <td class="bc-val-kemboja">${m.kVal}</td>
      <td class="bc-val-perdos">${m.pVal}</td>
      <td>${bcWinnerBadge(m.winner)}</td>
    </tr>`).join('');

  bcUpdateChart(r, k, p);
  bcSave(r, k, p);
  bcRenderTargets();
  buildHomeAlerts();
}

function bcUpdateChart(r, k, p) {
  if (typeof Chart === 'undefined') return;
  const isDark = document.body.getAttribute('data-theme') === 'dark';
  const textColor = isDark ? '#adb5bd' : '#6c757d';
  const gridColor = isDark ? 'rgba(255,255,255,0.06)' : 'rgba(0,0,0,0.06)';

  const labels = ['Revenue', 'Gross Profit', 'COGS'];
  const rayaData = [r.rev, r.rev - r.cogs, r.cogs];
  const kembData = [k.rev, k.rev - k.cogs, k.cogs];
  const perdData = [p.rev, p.rev - p.cogs, p.cogs];

  if (!bcChart) {
    bcChart = new Chart(document.getElementById('bcChart'), {
      type: 'bar',
      data: {
        labels,
        datasets: [
          { label: 'Raya',    data: rayaData, backgroundColor: 'rgba(99,102,241,0.8)',  borderRadius: 8 },
          { label: 'Kemboja', data: kembData, backgroundColor: 'rgba(245,158,11,0.8)',  borderRadius: 8 },
          { label: 'Perdos',  data: perdData, backgroundColor: 'rgba(16,185,129,0.8)',  borderRadius: 8 },
        ]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: {
          legend: { position: 'top', labels: { color: textColor, font: { family: 'Plus Jakarta Sans', weight: '700' }, padding: 16 } },
          datalabels: { display: false }
        },
        scales: {
          x: { ticks: { color: textColor, font: { weight: '700' } }, grid: { color: gridColor } },
          y: { beginAtZero: true, ticks: { color: textColor, callback: v => 'Rp' + formatShortNumber(v) }, grid: { color: gridColor } }
        }
      }
    });
  } else {
    bcChart.data.datasets[0].data = rayaData;
    bcChart.data.datasets[1].data = kembData;
    if (bcChart.data.datasets[2]) bcChart.data.datasets[2].data = perdData;
    bcChart.options.plugins.legend.labels.color = textColor;
    bcChart.options.scales.x.ticks.color = textColor;
    bcChart.options.scales.x.grid.color = gridColor;
    bcChart.options.scales.y.ticks.color = textColor;
    bcChart.options.scales.y.grid.color = gridColor;
    bcChart.update();
  }
}


// ═══════════════════════════════════════════
// RICH TOAST (replaces xlToast for new features)
// ═══════════════════════════════════════════
function richToast(msg, duration = 3500) {
  const el = document.getElementById('richToast');
  el.textContent = msg;
  el.classList.add('show');
  clearTimeout(el._t);
  el._t = setTimeout(() => el.classList.remove('show'), duration);
}

// ═══════════════════════════════════════════
// ONBOARDING CHECK
// ═══════════════════════════════════════════
function checkOnboarding() {
  const hasBranch = !!localStorage.getItem(LS_KEY_BRANCH);
  const hasRev = !!localStorage.getItem(LS_KEY_REV);
  const ob = document.getElementById('onboarding-banner');
  if (!hasBranch && !hasRev) { ob.classList.add('visible'); }
  else { ob.classList.remove('visible'); }
}

// ═══════════════════════════════════════════
// HOME ALERT BANNERS — auto-generate from stored data
// ═══════════════════════════════════════════
function buildHomeAlerts() {
  const wrap = document.getElementById('home-alert-wrap');
  if (!wrap) return;
  wrap.innerHTML = '';
  const alerts = [];

  // Check branch GM
  try {
    const bd = JSON.parse(localStorage.getItem(LS_KEY_BRANCH) || '{}');
    const branches = [
      { key: 'raya',  label: 'Raya',    color: '#6366f1' },
      { key: 'kemb',  label: 'Kemboja', color: '#f59e0b' },
      { key: 'perd',  label: 'Perdos',  color: '#10b981' },
    ];
    branches.forEach(b => {
      const rev  = parseFloat((bd[`${b.key}_rev`]  || '0').replace(/\./g,'')) || 0;
      const cogs = parseFloat((bd[`${b.key}_cogs`] || '0').replace(/\./g,'')) || 0;
      if (rev > 0) {
        const gm = (rev - cogs) / rev;
        if (gm < 0.38) {
          alerts.push({ type: 'danger',  icon: '🔴', title: `${b.label}: Gross Margin Merah!`, msg: `GM ${(gm*100).toFixed(1)}% — di bawah ambang batas 38%.` });
        } else if (gm < 0.45) {
          alerts.push({ type: 'warning', icon: '⚠️', title: `${b.label}: Margin Menipis`, msg: `GM ${(gm*100).toFixed(1)}% — mendekati zona merah.` });
        }
      }
    });
    // Check target vs actual
    const td = JSON.parse(localStorage.getItem(LS_KEY_TARGETS) || '{}');
    branches.forEach(b => {
      const rev    = parseFloat((bd[`${b.key}_rev`]  || '0').replace(/\./g,'')) || 0;
      const target = parseFloat((td[`tgt_${b.key}`]  || '0').replace(/[^0-9]/g,'')) || 0;
      if (rev > 0 && target > 0) {
        const pct = rev / target;
        if (pct < 0.75) {
          alerts.push({ type: 'danger', icon: '📉', title: `${b.label}: Target Jauh`, msg: `Baru ${(pct*100).toFixed(0)}% dari target bulan ini.` });
        }
      }
    });
  } catch(e) {}

  if (alerts.length === 0) {
    // If data exists but all good
    const hasBranch = !!localStorage.getItem(LS_KEY_BRANCH);
    if (hasBranch) {
      alerts.push({ type: 'success', icon: '✅', title: 'Semua Cabang Normal', msg: 'Tidak ada alert kritis saat ini.' });
    }
  }

  alerts.forEach((a, i) => {
    const div = document.createElement('div');
    div.className = `alert-banner alert-${a.type}`;
    div.style.animationDelay = `${i * 0.05}s`;
    div.innerHTML = `<div class="alert-banner-icon">${a.icon}</div><div class="alert-banner-body"><b>${a.title}</b>${a.msg}</div>`;
    wrap.appendChild(div);
  });
}

// ═══════════════════════════════════════════
// TARGET VS AKTUAL
// ═══════════════════════════════════════════
const LS_KEY_TARGETS = 'taobun_targets_v1';

function bcFmtTarget(el) {
  let v = el.value.replace(/[^0-9]/g, '');
  if (v) el.value = 'Rp ' + parseInt(v).toLocaleString('id-ID');
  bcSaveTargets();
}

function bcSaveTargets() {
  const d = {
    tgt_raya: document.getElementById('tgt-raya')?.value || '0',
    tgt_kemb: document.getElementById('tgt-kemb')?.value || '0',
    tgt_perd: document.getElementById('tgt-perd')?.value || '0',
  };
  try { localStorage.setItem(LS_KEY_TARGETS, JSON.stringify(d)); } catch(e) {}
}

function bcLoadTargets() {
  try {
    const d = JSON.parse(localStorage.getItem(LS_KEY_TARGETS) || '{}');
    if (d.tgt_raya) document.getElementById('tgt-raya').value = d.tgt_raya;
    if (d.tgt_kemb) document.getElementById('tgt-kemb').value = d.tgt_kemb;
    if (d.tgt_perd) document.getElementById('tgt-perd').value = d.tgt_perd;
    bcRenderTargets();
  } catch(e) {}
}

function bcRenderTargets() {
  const grid = document.getElementById('tgt-progress-grid');
  if (!grid) return;

  const branches = [
    { key: 'raya', label: 'Raya',    color: '#6366f1', revId: 'br-raya-rev', tgtId: 'tgt-raya' },
    { key: 'kemb', label: 'Kemboja', color: '#f59e0b', revId: 'br-kemb-rev', tgtId: 'tgt-kemb' },
    { key: 'perd', label: 'Perdos',  color: '#10b981', revId: 'br-perd-rev', tgtId: 'tgt-perd' },
  ];

  grid.innerHTML = branches.map(b => {
    const rev    = cleanNumber(document.getElementById(b.revId)?.value || '0');
    const rawTgt = document.getElementById(b.tgtId)?.value || '0';
    const target = parseFloat(rawTgt.replace(/[^0-9]/g,'')) || 0;
    if (target === 0) return `<div class="target-progress-card"><div class="tpc-branch" style="color:${b.color}">${b.label}</div><div class="tpc-pct" style="color:${b.color}">—</div><div class="tpc-label">Belum ada target</div></div>`;
    const pct = Math.min(rev / target, 2);
    const pctDisplay = (pct * 100).toFixed(1);
    const status = pct >= 1 ? { label: '✅ Tercapai', color: '#10b981' } : pct >= 0.75 ? { label: '⚡ On Track', color: '#f59e0b' } : { label: '🔴 Off Track', color: '#ef4444' };
    const barPct = Math.min(pct * 100, 100);
    return `
      <div class="target-progress-card">
        <div class="tpc-branch" style="color:${b.color}">${b.label}</div>
        <div class="tpc-pct" style="color:${status.color}">${pctDisplay}%</div>
        <div class="tpc-bar-wrap"><div class="tpc-bar" style="width:${barPct}%;background:${status.color}"></div></div>
        <div class="tpc-label">${bcRp(rev)} / ${bcRp(target)}</div>
        <div class="tpc-status" style="color:${status.color}">${status.label}</div>
      </div>`;
  }).join('');
  bcSaveTargets();
}

// ═══════════════════════════════════════════
// HISTORICAL TREND
// ═══════════════════════════════════════════
const LS_KEY_HISTORY = 'taobun_history_v1';
let historyChart = null;
let historyMetric = 'rev';

function bcGetHistory() {
  try { return JSON.parse(localStorage.getItem(LS_KEY_HISTORY) || '[]'); } catch(e) { return []; }
}

function bcSaveSnapshot() {
  const r = {
    rev:    bcGet('br-raya-rev'),  cust: bcGet('br-raya-cust'),
    tx:     bcGet('br-raya-tx'),   cogs: bcGet('br-raya-cogs'),
  };
  const k = {
    rev:    bcGet('br-kemb-rev'),  cust: bcGet('br-kemb-cust'),
    tx:     bcGet('br-kemb-tx'),   cogs: bcGet('br-kemb-cogs'),
  };
  const p = {
    rev:    bcGet('br-perd-rev'),  cust: bcGet('br-perd-cust'),
    tx:     bcGet('br-perd-tx'),   cogs: bcGet('br-perd-cogs'),
  };

  if (r.rev === 0 && k.rev === 0 && p.rev === 0) {
    richToast('⚠️ Isi data cabang terlebih dahulu sebelum menyimpan.');
    return;
  }

  const month = document.getElementById('selectedMonthBranch')?.textContent || 'Unknown';
  const year  = document.getElementById('selectedYearBranch')?.textContent  || new Date().getFullYear();
  const period = `${month} ${year}`;

  const history = bcGetHistory();
  const existIdx = history.findIndex(h => h.period === period);

  const snapshot = {
    period,
    ts: Date.now(),
    raya:    { ...r, gm: r.rev > 0 ? (r.rev - r.cogs) / r.rev : 0 },
    kemboja: { ...k, gm: k.rev > 0 ? (k.rev - k.cogs) / k.rev : 0 },
    perdos:  { ...p, gm: p.rev > 0 ? (p.rev - p.cogs) / p.rev : 0 },
  };

  if (existIdx >= 0) {
    showConfirm({ icon: '🔄', title: `Update ${period}?`, msg: `Data ${period} sudah ada. Timpa dengan data saat ini?`, okLabel: 'Update' }, function() {
      history[existIdx] = snapshot;
      localStorage.setItem(LS_KEY_HISTORY, JSON.stringify(history));
      bcRenderHistory();
      richToast(`✅ Data ${period} berhasil diupdate di histori.`);
    });
  } else {
    history.push(snapshot);
    // Keep max 24 periods
    if (history.length > 24) history.shift();
    localStorage.setItem(LS_KEY_HISTORY, JSON.stringify(history));
    bcRenderHistory();
    richToast(`✅ Data ${period} berhasil disimpan ke histori!`);
  }
}

function bcHistoryTab(btn) {
  document.querySelectorAll('.history-tab').forEach(t => t.classList.remove('active'));
  btn.classList.add('active');
  historyMetric = btn.dataset.metric;
  bcRenderHistory();
}

function bcRenderHistory() {
  const history = bcGetHistory();
  const wrap = document.getElementById('history-list-wrap');
  const canvas = document.getElementById('historyChart');
  if (!wrap || !canvas) return;

  if (history.length === 0) {
    wrap.innerHTML = `<div class="history-empty"><div class="history-empty-ico">📭</div>Belum ada data historis tersimpan.<br>Klik "Simpan Periode Ini" setelah mengisi data cabang.</div>`;
    if (historyChart) { historyChart.destroy(); historyChart = null; }
    canvas.style.display = 'none';
    return;
  }

  canvas.style.display = 'block';

  const labels = history.map(h => h.period);
  const metricFn = {
    rev:  h => [h.raya.rev, h.kemboja.rev, h.perdos.rev],
    gm:   h => [h.raya.gm * 100, h.kemboja.gm * 100, h.perdos.gm * 100],
    cust: h => [h.raya.cust, h.kemboja.cust, h.perdos.cust],
    tx:   h => [h.raya.tx, h.kemboja.tx, h.perdos.tx],
  };
  const fn = metricFn[historyMetric] || metricFn.rev;
  const rayaData   = history.map(h => fn(h)[0]);
  const kembojaData = history.map(h => fn(h)[1]);
  const perdosData  = history.map(h => fn(h)[2]);

  const isDark = document.body.getAttribute('data-theme') === 'dark';
  const gridColor = isDark ? 'rgba(255,255,255,0.06)' : 'rgba(0,0,0,0.06)';
  const tickColor = isDark ? '#adb5bd' : '#6c757d';

  const yLabel = historyMetric === 'rev' ? v => 'Rp' + formatShortNumber(v)
               : historyMetric === 'gm'  ? v => v.toFixed(1) + '%'
               : v => v.toLocaleString('id-ID');

  const datasets = [
    { label: 'Raya',    data: rayaData,    borderColor: '#6366f1', backgroundColor: 'rgba(99,102,241,0.1)',   tension: 0.4, fill: false, pointRadius: 4, pointHoverRadius: 6 },
    { label: 'Kemboja', data: kembojaData, borderColor: '#f59e0b', backgroundColor: 'rgba(245,158,11,0.1)',   tension: 0.4, fill: false, pointRadius: 4, pointHoverRadius: 6 },
    { label: 'Perdos',  data: perdosData,  borderColor: '#10b981', backgroundColor: 'rgba(16,185,129,0.1)',   tension: 0.4, fill: false, pointRadius: 4, pointHoverRadius: 6 },
  ];

  if (historyChart) {
    historyChart.data.labels = labels;
    historyChart.data.datasets[0].data = rayaData;
    historyChart.data.datasets[1].data = kembojaData;
    historyChart.data.datasets[2].data = perdosData;
    historyChart.options.scales.y.ticks.callback = yLabel;
    historyChart.options.scales.y.grid.color = gridColor;
    historyChart.options.scales.x.grid.color = gridColor;
    historyChart.options.scales.y.ticks.color = tickColor;
    historyChart.options.scales.x.ticks.color = tickColor;
    historyChart.options.plugins.legend.labels.color = tickColor;
    historyChart.update();
  } else {
    historyChart = new Chart(canvas, {
      type: 'line',
      data: { labels, datasets },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: {
          legend: { position: 'top', labels: { color: tickColor, font: { family: 'Plus Jakarta Sans', weight: '700' }, padding: 12 } },
          datalabels: { display: false }
        },
        scales: {
          x: { ticks: { color: tickColor, font: { weight: '600' } }, grid: { color: gridColor } },
          y: { beginAtZero: false, ticks: { color: tickColor, callback: yLabel }, grid: { color: gridColor } }
        }
      }
    });
  }

  // Render list
  wrap.innerHTML = `<div class="history-list">${[...history].reverse().map(h => `
    <div class="history-entry">
      <span class="history-entry-period">${h.period}</span>
      <div class="history-entry-vals">
        <span style="color:#6366f1;font-size:11px;font-weight:700">${bcRp(h.raya.rev)}</span>
        <span style="color:#f59e0b;font-size:11px;font-weight:700">${bcRp(h.kemboja.rev)}</span>
        <span style="color:#10b981;font-size:11px;font-weight:700">${bcRp(h.perdos.rev)}</span>
        <button class="history-entry-del" onclick="bcDeleteHistory('${h.period}')" title="Hapus">✕</button>
      </div>
    </div>`).join('')}</div>`;
}

function bcDeleteHistory(period) {
  showConfirm({ icon: '🗑️', title: `Hapus ${period}?`, msg: 'Data histori periode ini akan dihapus permanen.', okLabel: 'Hapus' }, function() {
    const history = bcGetHistory().filter(h => h.period !== period);
    localStorage.setItem(LS_KEY_HISTORY, JSON.stringify(history));
    bcRenderHistory();
    richToast(`🗑️ Data ${period} dihapus dari histori.`);
  });
}

// ═══════════════════════════════════════════
// PDF EXPORT (print-based)
// ═══════════════════════════════════════════
function bcExportPDF() {
  const month = document.getElementById('selectedMonthBranch')?.textContent || '';
  const year  = document.getElementById('selectedYearBranch')?.textContent  || '';
  const period = `${month} ${year}`;

  const r = { rev: bcGet('br-raya-rev'), cogs: bcGet('br-raya-cogs'), cust: bcGet('br-raya-cust'), tx: bcGet('br-raya-tx') };
  const k = { rev: bcGet('br-kemb-rev'), cogs: bcGet('br-kemb-cogs'), cust: bcGet('br-kemb-cust'), tx: bcGet('br-kemb-tx') };
  const p = { rev: bcGet('br-perd-rev'), cogs: bcGet('br-perd-cogs'), cust: bcGet('br-perd-cust'), tx: bcGet('br-perd-tx') };

  r.gm = r.rev > 0 ? ((r.rev - r.cogs) / r.rev * 100).toFixed(1) : '—';
  k.gm = k.rev > 0 ? ((k.rev - k.cogs) / k.rev * 100).toFixed(1) : '—';
  p.gm = p.rev > 0 ? ((p.rev - p.cogs) / p.rev * 100).toFixed(1) : '—';

  const totalRev = r.rev + k.rev + p.rev;

  const tgt = JSON.parse(localStorage.getItem(LS_KEY_TARGETS) || '{}');
  const tgtRaya = parseFloat((tgt.tgt_raya || '0').replace(/[^0-9]/g, '')) || 0;
  const tgtKemb = parseFloat((tgt.tgt_kemb || '0').replace(/[^0-9]/g, '')) || 0;
  const tgtPerd = parseFloat((tgt.tgt_perd || '0').replace(/[^0-9]/g, '')) || 0;

  const pctRow = (rev, tgt) => tgt > 0 ? `${(rev/tgt*100).toFixed(1)}%` : '—';

  const isDark = document.body.getAttribute('data-theme') === 'dark';
  const html = `<!DOCTYPE html><html><head><meta charset="UTF-8">
<title>Ringkasan Branch — ${period}</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Segoe UI', Arial, sans-serif; background: #f8f9fa; color: #1a1a2e; padding: 32px; font-size: 13px; }
  h1 { font-size: 22px; font-weight: 800; margin-bottom: 4px; }
  .sub { color: #6c757d; font-size: 12px; margin-bottom: 24px; }
  .logo { color: #d61e30; }
  table { width: 100%; border-collapse: collapse; margin-bottom: 24px; background: white; border-radius: 12px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }
  th { background: #f8f9fa; padding: 10px 14px; text-align: left; font-size: 10px; font-weight: 800; text-transform: uppercase; letter-spacing: .7px; color: #6c757d; }
  td { padding: 12px 14px; border-top: 1px solid #e9ecef; font-size: 12.5px; }
  tr:last-child td { font-weight: 800; }
  .raya { color: #6366f1; } .kemb { color: #f59e0b; } .perd { color: #10b981; }
  .total { background: #d61e30; color: white; }
  .total th { background: rgba(255,255,255,0.15); color: white; }
  .total td { color: white; border-color: rgba(255,255,255,0.15); }
  .section-title { font-size: 11px; font-weight: 800; text-transform: uppercase; letter-spacing: .8px; color: #6c757d; margin-bottom: 10px; border-left: 3px solid #d61e30; padding-left: 10px; }
  .footer { margin-top: 28px; font-size: 10px; color: #adb5bd; border-top: 1px solid #e9ecef; padding-top: 12px; }
</style>
</head><body>
<h1><span class="logo">TAOBUN</span> Branch Report</h1>
<div class="sub">Periode: ${period} &nbsp;•&nbsp; Dicetak: ${new Date().toLocaleDateString('id-ID', {weekday:'long', year:'numeric', month:'long', day:'numeric'})}</div>

<div class="section-title">Revenue Summary</div>
<table>
  <thead><tr><th>Cabang</th><th>Revenue</th><th>COGS</th><th>Gross Profit</th><th>GM %</th><th>Customers</th><th>Transaksi</th></tr></thead>
  <tbody>
    <tr><td class="raya"><b>🟣 Raya</b></td><td>${bcRp(r.rev)}</td><td>${bcRp(r.cogs)}</td><td>${bcRp(r.rev-r.cogs)}</td><td>${r.gm}%</td><td>${r.cust.toLocaleString('id-ID')}</td><td>${r.tx.toLocaleString('id-ID')}</td></tr>
    <tr><td class="kemb"><b>🟡 Kemboja</b></td><td>${bcRp(k.rev)}</td><td>${bcRp(k.cogs)}</td><td>${bcRp(k.rev-k.cogs)}</td><td>${k.gm}%</td><td>${k.cust.toLocaleString('id-ID')}</td><td>${k.tx.toLocaleString('id-ID')}</td></tr>
    <tr><td class="perd"><b>🟢 Perdos</b></td><td>${bcRp(p.rev)}</td><td>${bcRp(p.cogs)}</td><td>${bcRp(p.rev-p.cogs)}</td><td>${p.gm}%</td><td>${p.cust.toLocaleString('id-ID')}</td><td>${p.tx.toLocaleString('id-ID')}</td></tr>
    <tr><td><b>TOTAL</b></td><td><b>${bcRp(totalRev)}</b></td><td><b>${bcRp(r.cogs+k.cogs+p.cogs)}</b></td><td><b>${bcRp(totalRev - r.cogs - k.cogs - p.cogs)}</b></td><td>—</td><td><b>${(r.cust+k.cust+p.cust).toLocaleString('id-ID')}</b></td><td><b>${(r.tx+k.tx+p.tx).toLocaleString('id-ID')}</b></td></tr>
  </tbody>
</table>

${(tgtRaya||tgtKemb||tgtPerd) ? `
<div class="section-title">Target vs Aktual</div>
<table>
  <thead><tr><th>Cabang</th><th>Target</th><th>Aktual</th><th>Pencapaian</th><th>Status</th></tr></thead>
  <tbody>
    <tr><td class="raya"><b>🟣 Raya</b></td><td>${tgtRaya?bcRp(tgtRaya):'—'}</td><td>${bcRp(r.rev)}</td><td>${pctRow(r.rev,tgtRaya)}</td><td>${tgtRaya?(r.rev>=tgtRaya?'✅ Tercapai':r.rev/tgtRaya>=0.75?'⚡ On Track':'🔴 Off Track'):'—'}</td></tr>
    <tr><td class="kemb"><b>🟡 Kemboja</b></td><td>${tgtKemb?bcRp(tgtKemb):'—'}</td><td>${bcRp(k.rev)}</td><td>${pctRow(k.rev,tgtKemb)}</td><td>${tgtKemb?(k.rev>=tgtKemb?'✅ Tercapai':k.rev/tgtKemb>=0.75?'⚡ On Track':'🔴 Off Track'):'—'}</td></tr>
    <tr><td class="perd"><b>🟢 Perdos</b></td><td>${tgtPerd?bcRp(tgtPerd):'—'}</td><td>${bcRp(p.rev)}</td><td>${pctRow(p.rev,tgtPerd)}</td><td>${tgtPerd?(p.rev>=tgtPerd?'✅ Tercapai':p.rev/tgtPerd>=0.75?'⚡ On Track':'🔴 Off Track'):'—'}</td></tr>
  </tbody>
</table>` : ''}

<div class="footer">TAOBUN Management Suite • CV. Roti Semua Generasi • Laporan ini di-generate otomatis dari sistem.</div>
</body></html>`;

  const win = window.open('', '_blank');
  win.document.write(html);
  win.document.close();
  setTimeout(() => { win.focus(); win.print(); }, 400);
  richToast(`📄 Membuka preview PDF untuk ${period}...`);
}

function bcSave(r, k, p) {
  const data = {};
  const fields = ['rev','cust','tx','basket','new_c','cogs'];
  fields.forEach(f => {
    const sfx = f === 'new_c' ? 'new' : f;
    data[`raya_${f}`] = document.getElementById(`br-raya-${sfx}`)?.value || '0';
    data[`kemb_${f}`] = document.getElementById(`br-kemb-${sfx}`)?.value || '0';
    data[`perd_${f}`] = document.getElementById(`br-perd-${sfx}`)?.value || '0';
  });
  try { localStorage.setItem(LS_KEY_BRANCH, JSON.stringify(data)); } catch(e) {}
}

function bcLoad() {
  try {
    const raw = localStorage.getItem(LS_KEY_BRANCH);
    if (!raw) return;
    const d = JSON.parse(raw);
    const map = { rev: 'rev', cust: 'cust', tx: 'tx', basket: 'basket', new_c: 'new', cogs: 'cogs' };
    Object.entries(map).forEach(([key, sfx]) => {
      const rEl = document.getElementById(`br-raya-${sfx}`);
      const kEl = document.getElementById(`br-kemb-${sfx}`);
      const pEl = document.getElementById(`br-perd-${sfx}`);
      if (rEl && d[`raya_${key}`]) rEl.value = d[`raya_${key}`];
      if (kEl && d[`kemb_${key}`]) kEl.value = d[`kemb_${key}`];
      if (pEl && d[`perd_${key}`]) pEl.value = d[`perd_${key}`];
    });
    bcCalculate();
    bcLoadTargets();
    bcRenderHistory();
    updatePeriodBadge();
  } catch(e) {}
}

function bcReset() {
  showConfirm({
    icon: '🗑️', title: 'Reset Branch Data?',
    msg: 'Semua data input cabang akan dikosongkan. Target dan histori tetap tersimpan.',
    okLabel: 'Reset'
  }, function() {
    ['rev','cust','tx','basket','new','cogs'].forEach(f => {
      ['raya','kemb','perd'].forEach(b => {
        const el = document.getElementById(`br-${b}-${f}`);
        if (el) el.value = '0';
      });
    });
    try { localStorage.removeItem(LS_KEY_BRANCH); } catch(e) {}
    bcCalculate();
    richToast('✅ Data ketiga cabang berhasil direset.');
    buildHomeAlerts();
    checkOnboarding();
  });
}

// ═══════════════════════════════════════════
// Init sobun on load
// ═══════════════════════════════════════════
// ═══════════════════════════════════════════
// PASSWORD GATE — Sobat Taobun
// ═══════════════════════════════════════════
// Password disimpan sebagai hash SHA-256.
// Untuk ganti password: jalankan fungsi pwHashPassword('passwordbaru')
// di console browser, lalu simpan hash-nya ke PW_HASH_KEY di localStorage.
// ═══════════════════════════════════════════
const PW_HASH_KEY = 'taobun_pw_hash_v1';
// Hash SHA-256 dari password default (taobun2026).
// Ganti string ini jika ingin ubah password default.
const PW_DEFAULT_HASH = 'e53ceb6dfdb5845b305c826e2da049f495d66037df60056d14ee8bcea7779900';

async function pwHashPassword(plain) {
  const enc = new TextEncoder().encode(plain);
  const buf = await crypto.subtle.digest('SHA-256', enc);
  return Array.from(new Uint8Array(buf)).map(b => b.toString(16).padStart(2,'0')).join('');
}

async function pwGetHash() {
  try { return localStorage.getItem(PW_HASH_KEY) || PW_DEFAULT_HASH; } catch(e) { return PW_DEFAULT_HASH; }
}

function pwOpen() {
  const overlay = document.getElementById('pwOverlay');
  overlay.classList.add('open');
  const inp = document.getElementById('pwInput');
  inp.value = '';
  inp.type = 'password';
  document.getElementById('pwToggleBtn').textContent = '👁';
  document.getElementById('pwErrMsg').textContent = '';
  inp.classList.remove('pw-error');
  setTimeout(() => inp.focus(), 150);
}

function pwClose() {
  document.getElementById('pwOverlay').classList.remove('open');
}

function pwToggleShow() {
  const inp = document.getElementById('pwInput');
  const btn = document.getElementById('pwToggleBtn');
  if (inp.type === 'password') { inp.type = 'text'; btn.textContent = '🙈'; }
  else { inp.type = 'password'; btn.textContent = '👁'; }
}

async function pwSubmit() {
  const inp = document.getElementById('pwInput');
  const val = inp.value;
  if (!val) {
    document.getElementById('pwErrMsg').textContent = 'Password tidak boleh kosong.';
    inp.classList.add('pw-error');
    return;
  }
  const inputHash = await pwHashPassword(val);
  const storedHash = await pwGetHash();
  if (inputHash !== storedHash) {
    document.getElementById('pwErrMsg').textContent = '❌ Password salah. Coba lagi.';
    inp.classList.add('pw-error');
    inp.value = '';
    setTimeout(() => inp.classList.remove('pw-error'), 700);
    return;
  }
  window._sobunUnlocked = true;
  sessionStorage.setItem('taobun_sobun_unlocked', '1');
  pwClose();
  showView('sobun-view');
}

// ── Import Excel password gate ──
function pwImportClose() {
  document.getElementById('pwImportOverlay').classList.remove('open');
  xlParsed = null;
  xlToast('ℹ️ Update data dibatalkan.');
}

function pwImportToggleShow() {
  const inp = document.getElementById('pwImportInput');
  const btn = document.getElementById('pwImportToggleBtn');
  if (inp.type === 'password') { inp.type = 'text'; btn.textContent = '🙈'; }
  else { inp.type = 'password'; btn.textContent = '👁'; }
}

async function pwImportSubmit() {
  const inp = document.getElementById('pwImportInput');
  const val = inp.value;
  if (!val) {
    document.getElementById('pwImportErrMsg').textContent = 'Password tidak boleh kosong.';
    inp.classList.add('pw-error');
    return;
  }
  const inputHash = await pwHashPassword(val);
  const storedHash = await pwGetHash();
  if (inputHash !== storedHash) {
    document.getElementById('pwImportErrMsg').textContent = '❌ Password salah. Coba lagi.';
    inp.classList.add('pw-error');
    inp.value = '';
    setTimeout(() => inp.classList.remove('pw-error'), 700);
    return;
  }
  // Password benar — terapkan semua data Excel
  document.getElementById('pwImportOverlay').classList.remove('open');
  const d = xlParsed;
  if (!d) { xlToast('⚠️ Tidak ada data Excel yang menunggu.'); return; }

  const changes = [];
  const rp = v => 'Rp ' + Math.round(v).toLocaleString('id-ID');

  // Sheet 1: Master Produk
  if (d.products && d.products.length > 0) {
    const oldCount = P.length;
    P.length = 0; d.products.forEach(p => P.push(p));
    changes.push(`📦 Master Produk: ${oldCount} → ${d.products.length} produk`);
  }

  // Sheet 2: Skema Tiering
  if (d.tiering) {
    const setRp = (id, v) => { const el = document.getElementById(id); if(el) el.value = 'Rp ' + v.toLocaleString('id-ID'); };
    setRp('t-bp', d.tiering.bp); setRp('p-bp', d.tiering.bp);
    setRp('t-np', d.tiering.np); setRp('p-np', d.tiering.np);
    changes.push(`⚙️ Tiering: ${rp(d.tiering.bp)}/poin, nilai ${rp(d.tiering.np)}`);
  }

  // Sheet 3: Simulasi Profitabilitas — HPP % + avg struk per tier
  if (d.hppPct !== undefined) {
    const el = document.getElementById('p-hpp');
    if(el) el.value = d.hppPct.toFixed(2);
    changes.push(`📊 HPP Assumption: ${d.hppPct.toFixed(2)}%`);
  }
  if (d.avgStruk) {
    const setRp2 = (id, v) => { const el = document.getElementById(id); if(el && v > 0) el.value = 'Rp ' + v.toLocaleString('id-ID'); };
    setRp2('p-abr', d.avgStruk.bronze);
    setRp2('p-asi', d.avgStruk.silver);
    setRp2('p-ago', d.avgStruk.gold);
    setRp2('p-apl', d.avgStruk.platinum);
    const filled = ['bronze','silver','gold','platinum'].filter(k=>d.avgStruk[k]>0);
    if(filled.length) changes.push(`📊 Avg Struk: ${filled.length} tier diupdate`);
  }

  // Sheet 3: Tier thresholds (naik/stay) — update TIER_TH state & display
  if (d.thresholds) {
    let tCount = 0;
    // Map dari Excel: silver naik=3jt→250rb? 
    // Excel pakai nilai TAHUNAN (3.000.000). Di UI pakai shorthand.
    // Tapi nilai di Excel row14-16 col2=NAIK col3=STAY
    // Silver: naik=3000000 di Excel = threshold naik dari bronze ke silver
    // Kita map: bronze.naik = silver threshold NAIK (belanja min utk naik dari bronze)
    const th = d.thresholds;
    if (th.silver)   { TIER_TH.bronze.naik  = th.silver.naik;   TIER_TH.silver.stay  = th.silver.stay;   tCount += 2; }
    if (th.gold)     { TIER_TH.silver.naik  = th.gold.naik;     TIER_TH.gold.stay    = th.gold.stay;     tCount += 2; }
    if (th.platinum) { TIER_TH.gold.naik    = th.platinum.naik; TIER_TH.platinum.stay= th.platinum.stay; tCount += 2; }
    TIER_TH.platinum.naik = TIER_TH.platinum.stay;
    updateTierDisplay();
    if (tCount) changes.push('🎯 Threshold Tier: diupdate dari Excel');
  }

  // Sheet 4: Simulasi Misi & Promo
  if (d.promos && d.promos.length > 0) {
    promos = d.promos;
    renderPromo();
    const aman = d.promos.filter(p=>p.gm>=0.45).length;
    const boncos = d.promos.filter(p=>p.gm<0.38).length;
    changes.push(`🎯 Promo: ${d.promos.length} program (${aman} aman${boncos?' · '+boncos+' boncos':''})`);
  }

  // Sheet 5: Simulasi Reward
  if (d.rewards && d.rewards.length > 0) {
    RW = d.rewards;
    renderRw();
    const byTier = {bronze:0,silver:0,gold:0,platinum:0};
    d.rewards.forEach(r=>{ if(byTier[r.t]!==undefined) byTier[r.t]++; });
    const summary = Object.entries(byTier).filter(([,v])=>v>0).map(([k,v])=>v+' '+k).join(', ');
    changes.push(`🏆 Reward: ${d.rewards.length} reward (${summary})`);
  }

  initDash(); renderMaster(); calcT(); calcP(); renderRw(); cbRender();
  sbSave(); // save AFTER apply, before any load
  
  // ── SIMPAN SEBAGAI EXCEL BASELINE (data permanen) ──
  // Data ini akan jadi default yang tidak bisa diubah sampai Excel baru diimport
  try {
    const baseline = {
      importedAt: new Date().toISOString(),
      products: d.products && d.products.length > 0 ? d.products.slice() : null,
      tiering: d.tiering || null,
      hppPct: d.hppPct !== undefined ? d.hppPct : null,
      avgStruk: d.avgStruk || null,
      thresholds: d.thresholds || null,
      promos: d.promos && d.promos.length > 0 ? d.promos.slice() : null,
      rewards: d.rewards && d.rewards.length > 0 ? d.rewards.slice() : null,
      // Snapshot nilai field saat ini untuk restore
      fields: {
        t_bp: document.getElementById('t-bp')?.value,
        t_np: document.getElementById('t-np')?.value,
        p_hpp: document.getElementById('p-hpp')?.value,
        p_bp: document.getElementById('p-bp')?.value,
        p_np: document.getElementById('p-np')?.value,
        p_abr: document.getElementById('p-abr')?.value,
        p_asi: document.getElementById('p-asi')?.value,
        p_ago: document.getElementById('p-ago')?.value,
        p_apl: document.getElementById('p-apl')?.value,
      },
      tierThresholds: JSON.parse(JSON.stringify(TIER_TH)),
    };
    localStorage.setItem(LS_KEY_EXCEL_BASELINE, JSON.stringify(baseline));
    _xlBaselineLoaded = true;
    applyExcelBaslineLock(true);
    console.log('[Excel Baseline] Disimpan:', new Date(baseline.importedAt).toLocaleString('id-ID'));
  } catch(e) { console.warn('[Excel Baseline] Gagal simpan:', e); }
  
  xlParsed = null;

  const msg = changes.length
    ? '✅ Data diupdate!\n' + changes.join('\n')
    : '✅ Data diupdate & disimpan!';
  xlToast(msg.split('\n')[0] + (changes.length > 1 ? ` (+${changes.length-1} lainnya)` : '') + ' · 🔒 Disimpan sebagai data default');
  console.log('[Import Sobat Taobun] Perubahan:\n' + changes.join('\n'));
}

function initSobun(){
  initDash();renderMaster();calcT();calcP();renderRw();cbRender();sbLoad();
  renderPromo(); // Pastikan promo dari baseline ikut dirender
}

window.onload = () => {
    // Sync anti-flash theme from <html> to <body> immediately
    const savedTheme = document.documentElement.getAttribute('data-theme');
    if (savedTheme) document.body.setAttribute('data-theme', savedTheme);
    else document.body.setAttribute('data-theme', 'light');

    // Build month/year dropdowns
    const monthContainers = document.querySelectorAll('.month-options-list');
    monthContainers.forEach(container => {
        months.forEach(m => {
            let d = document.createElement('div'); d.className = 'option-item'; d.textContent = m.l;
            d.onclick = () => selectOption('Month', m.v, m.l);
            container.appendChild(d);
        });
    });

    const yearContainers = document.querySelectorAll('.year-options-list');
    yearContainers.forEach(container => {
        for(let i=2026; i<=2030; i++) { 
            let d = document.createElement('div'); d.className = 'option-item'; d.textContent = i;
            d.onclick = () => selectOption('Year', i.toString(), i.toString());
            container.appendChild(d);
        }
    });

    // Init core dashboard first (fast), defer heavy Sobun init slightly
    initCharts();
    loadData();
    // Apply correct chart colors after theme is loaded from localStorage
    updateChartTheme();

    // Keybind logic
    const inputs = Array.from(document.querySelectorAll('.nav-input, .lead-field'));
    inputs.forEach((input, index) => {
        input.addEventListener('keydown', (e) => {
            if (e.key === 'Enter' || e.key === 'ArrowDown') { e.preventDefault(); if (inputs[index+1]) inputs[index+1].focus(); }
            else if (e.key === 'ArrowUp') { e.preventDefault(); if (inputs[index-1]) inputs[index-1].focus(); }
        });
    });

    // Defer Sobun init to next frame so home view renders first
    requestAnimationFrame(() => {
        initSobun();
        bcLoad();
        updatePeriodBadge();
        checkOnboarding();
        buildHomeAlerts();

        // ── RESTORE VIEW AKTIF setelah refresh ──
        try {
            const savedView = localStorage.getItem('taobun_active_view');
            if (savedView && savedView !== 'home-view') {
                const targetEl = document.getElementById(savedView);
                if (targetEl) {
                    if (savedView === 'sobun-view') {
                        // Cek apakah sobun sudah pernah di-unlock di session ini
                        // atau di session sebelumnya (simpan di sessionStorage)
                        const wasUnlocked = sessionStorage.getItem('taobun_sobun_unlocked') === '1';
                        if (wasUnlocked) {
                            window._sobunUnlocked = true;
                            // Langsung tampilkan sobun tanpa password gate
                            document.querySelectorAll('.view-container').forEach(v => v.classList.remove('view-active'));
                            targetEl.classList.add('view-active');
                        }
                        // Kalau belum unlock, biarkan di home
                    } else {
                        showView(savedView);
                    }
                }
            }
        } catch(e) {}

        // ── RESTORE TAB AKTIF SOBAT TAOBUN setelah refresh ──
        try {
            const savedTab = localStorage.getItem('taobun_active_sobun_tab');
            if (savedTab) {
                const tabEl = document.getElementById('sb-page-' + savedTab);
                const navBtn = Array.from(document.querySelectorAll('.sobun-wrapper .nav-tab'))
                    .find(btn => btn.getAttribute('onclick') && btn.getAttribute('onclick').includes("'" + savedTab + "'"));
                if (tabEl && navBtn) {
                    document.querySelectorAll('.sobun-wrapper .page').forEach(p => p.classList.remove('active'));
                    document.querySelectorAll('.sobun-wrapper .nav-tab').forEach(t => t.classList.remove('active'));
                    tabEl.classList.add('active');
                    navBtn.classList.add('active');
                    if (savedTab === 'master')         renderMaster();
                    else if (savedTab === 'reward')    renderRw();
                    else if (savedTab === 'dashboard') initDash();
                    else if (savedTab === 'tiering')   calcT();
                    else if (savedTab === 'profit')    calcP();
                    else if (savedTab === 'promo')     cbRender();
                }
            }
        } catch(e) {}
    });

    // Close dropdowns on outside click
    document.addEventListener('click', (e) => {
        if (!e.target.closest('.custom-select-container')) {
            document.querySelectorAll('.custom-select-container').forEach(c => c.classList.remove('active'));
        }
    });
};
