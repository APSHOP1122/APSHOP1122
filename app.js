// ============================================================================
// AP COSMETICS (ร้านเครื่องสำอาง AP) — ระบบบริหารคลังสินค้าและจุดขาย (POS)
// Organization: APSHOP1122'sOrg
// Supabase URL: https://brvvogwestgptyfohrba.supabase.co
// ============================================================================

const SUPABASE_URL = 'https://brvvogwestgptyfohrba.supabase.co';
const SUPABASE_KEY = 'sb_publishable_iZFnhJIwesIkI8jSsTA0Rw_Zx_tITr-';
const db = window.supabase.createClient(SUPABASE_URL, SUPABASE_KEY);

// Default Excel Reference Prices (from Inventory_TB.xlsx)
const EXCEL_REFERENCE_ITEMS = [
  { In_ID: 1, In_Bar: 'PA001', In_Menu: 'บิ๊กอาย', In_Unit: 'คู่', cost: 40, sell: 100 },
  { In_ID: 2, In_Bar: 'PA002', In_Menu: 'ต่างหู', In_Unit: 'คู่', cost: 10, sell: 30 },
  { In_ID: 3, In_Bar: 'PA003', In_Menu: 'สีย้อมผม Focus', In_Unit: 'กล่อง', cost: 35, sell: 59 },
  { In_ID: 4, In_Bar: 'PA004', In_Menu: 'สีย้อมผม Gosen', In_Unit: 'กล่อง', cost: 75, sell: 90 },
  { In_ID: 5, In_Bar: 'PA005', In_Menu: 'สีย้อมผม Farger', In_Unit: 'กล่อง', cost: 55, sell: 79 },
  { In_ID: 6, In_Bar: 'PA006', In_Menu: 'ทรีทเม้นท์เปลี่ยนสีผม', In_Unit: 'กล่อง', cost: 90, sell: 129 },
  { In_ID: 7, In_Bar: 'PA007', In_Menu: 'ครีมเปลี่ยนสีผม Schwarzkopf', In_Unit: 'ซอง', cost: 270, sell: 299 },
  { In_ID: 8, In_Bar: 'PA008', In_Menu: 'ครีมเปลี่ยนสีผม Esane', In_Unit: 'ซอง', cost: 90, sell: 129 },
  { In_ID: 9, In_Bar: 'PA009', In_Menu: 'ครีมเปลี่ยนสีผมเบอริน่า', In_Unit: 'ซอง', cost: 38, sell: 59 },
  { In_ID: 10, In_Bar: 'PA010', In_Menu: 'ดินสอเขียนคิ้ว', In_Unit: 'แท่ง', cost: 5, sell: 10 },
  { In_ID: 11, In_Bar: 'PA011', In_Menu: 'เกลือขัดผิว', In_Unit: 'ซอง', cost: 15, sell: 29 },
  { In_ID: 12, In_Bar: 'PA012', In_Menu: 'มาร์คผม Anglewink', In_Unit: 'กระปุก', cost: 180, sell: 250 },
  { In_ID: 13, In_Bar: 'PA013', In_Menu: 'แชมพู Angle', In_Unit: 'ขวด', cost: 150, sell: 250 },
  { In_ID: 14, In_Bar: 'PA014', In_Menu: 'สเปรย์ฟื้นฟูผม', In_Unit: 'ขวด', cost: 180, sell: 250 },
  { In_ID: 15, In_Bar: 'PA015', In_Menu: 'ฟองน้ำเช็ดหน้า', In_Unit: 'อัน', cost: 15, sell: 30 },
  { In_ID: 16, In_Bar: 'PA016', In_Menu: 'Olay ครีมบำรุง', In_Unit: 'กระปุก', cost: 100, sell: 129 },
  { In_ID: 17, In_Bar: 'PA017', In_Menu: 'ครีมทาขาหนีบ', In_Unit: 'กระปุก', cost: 70, sell: 150 },
  { In_ID: 18, In_Bar: 'PA018', In_Menu: 'ผงหอมศรีจันทร์ทานาคา', In_Unit: 'ตลับ', cost: 80, sell: 129 },
  { In_ID: 19, In_Bar: 'PA019', In_Menu: 'เทปติดขนตาสองชั้น', In_Unit: 'แผง', cost: 45, sell: 99 },
  { In_ID: 20, In_Bar: 'PA020', In_Menu: 'LC ทรีทเม้นท์ผม', In_Unit: 'กระปุก', cost: 100, sell: 150 },
  { In_ID: 21, In_Bar: 'PA021', In_Menu: 'แปรงปัดคิ้ว', In_Unit: 'อัน', cost: 20, sell: 39 },
  { In_ID: 22, In_Bar: 'PA022', In_Menu: 'หวีรีดผม', In_Unit: 'อัน', cost: 60, sell: 120 },
  { In_ID: 23, In_Bar: 'PA023', In_Menu: 'กรรไกรตัดขนจมูก', In_Unit: 'อัน', cost: 20, sell: 39 },
  { In_ID: 24, In_Bar: 'PA024', In_Menu: 'สีทาเล็บ', In_Unit: 'ขวด', cost: 20, sell: 29 },
  { In_ID: 25, In_Bar: 'PA025', In_Menu: 'เล็บปลอม', In_Unit: 'กล่อง', cost: 10, sell: 29 },
  { In_ID: 26, In_Bar: 'PA026', In_Menu: 'มายนา ซันสกรีน', In_Unit: 'ขวด', cost: 100, sell: 199 },
  { In_ID: 27, In_Bar: 'PA027', In_Menu: 'ครีมบำรุงเส้นผม Wing', In_Unit: 'ซอง', cost: 15, sell: 29 }
];

const EXCEL_MAP = {};
EXCEL_REFERENCE_ITEMS.forEach(item => {
  EXCEL_MAP[item.In_Bar] = item;
});

// Application State
let currentUser = null; // { role: 'admin' | 'user', name, email, code }
let inventoryList = [];
let purchaseSelectList = []; // From SELECT In_Bar, In_Menu, In_Price
let purchaseSummaryList = [];
let salesSummaryList = [];
let expenseSummaryList = [];
let memberList = [];

let purchaseCart = [];
let salesCart = [];

// ============================================================================
// UTILITY & TOAST HELPERS
// ============================================================================
function showToast(message, type = 'info') {
  const container = document.getElementById('toastContainer');
  if (!container) return;
  const el = document.createElement('div');
  el.className = `toast ${type}`;
  el.innerHTML = message;
  container.appendChild(el);
  setTimeout(() => {
    el.style.opacity = '0';
    setTimeout(() => el.remove(), 300);
  }, 3800);
}

function getTodayStr() {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, '0');
  const d = String(now.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

function formatMoney(num) {
  return Number(num || 0).toLocaleString('th-TH', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function getUnitCostPrice(row) {
  const ref = EXCEL_MAP[row.In_Bar];
  if (row.In_Price !== undefined && row.In_Price !== null && Number(row.In_Price) > 0) return Number(row.In_Price);
  if (row.In_PriceSum !== undefined && row.In_PriceSum !== null && Number(row.In_PriceSum) > 0) return Number(row.In_PriceSum);
  return ref ? ref.cost : 0;
}

function getUnitSellPrice(row) {
  const ref = EXCEL_MAP[row.In_Bar];
  if (row.In_Sell_Price !== undefined && row.In_Sell_Price !== null && Number(row.In_Sell_Price) > 0) return Number(row.In_Sell_Price);
  return ref ? ref.sell : 0;
}

// ============================================================================
// ASTERISK MASKING FOR ADMIN USERNAME & PASSWORD FIELDS
// ("แอดมินห้ามโชว์ชื่อผู้ใช้งานและรหัสผู้ใช้งานทั้งสองช่องนี้ให้ทำเป็นอักขระเครื่องหมายดอกจันทร์")
// ============================================================================
function maskAsteriskField(visibleInput, hiddenId) {
  const hiddenEl = document.getElementById(hiddenId);
  if (!hiddenEl) return;

  const currentReal = hiddenEl.value || '';
  const rawVal = visibleInput.value || '';

  if (rawVal.length === 0) {
    hiddenEl.value = '';
    visibleInput.value = '';
    return;
  }

  // Reconstruct actual value from non-asterisk chars typed or backspace
  let newReal = '';
  if (rawVal.length < currentReal.length && /^\*+$/.test(rawVal)) {
    newReal = currentReal.slice(0, rawVal.length);
  } else {
    let idx = 0;
    for (let i = 0; i < rawVal.length; i++) {
      if (rawVal[i] === '*') {
        newReal += currentReal[idx] || '';
        idx++;
      } else {
        newReal += rawVal[i];
      }
    }
  }

  hiddenEl.value = newReal;
  visibleInput.value = '*'.repeat(newReal.length);
}

// ============================================================================
// AUTHENTICATION & ROLE VISIBILITY
// ============================================================================
function handleAdminLogin(e) {
  e.preventDefault();
  const user = (document.getElementById('adminUserReal').value || '').trim();
  const pass = (document.getElementById('adminPassReal').value || '').trim();

  if (user.toUpperCase() === 'ADMIN' && pass === '1122') {
    currentUser = {
      role: 'admin',
      name: 'ผู้ดูแลระบบสูงสุด (ADMIN)',
      email: 'admin@apcosmetics.co.th',
      code: 'ADMIN-1122'
    };
    enterApplication();
    showToast('เข้าสู่ระบบแอดมินสำเร็จ! เปิดใช้งานครบทุกเมนูและระบบจัดการสมาชิก', 'success');
  } else {
    showToast('ชื่อผู้ใช้งานหรือรหัสผ่านแอดมินไม่ถูกต้อง', 'error');
  }
}

function autoFillUserCodeByEmail(emailVal) {
  const cleanEmail = (emailVal || '').trim().toLowerCase();
  const found = memberList.find(m => (m.EM_Email || '').toLowerCase() === cleanEmail);
  if (found) {
    document.getElementById('userLoginCode').value = found.Em_ID;
  }
}

function selectQuickUser(emId) {
  if (!emId) return;
  const found = memberList.find(m => m.Em_ID === emId);
  if (found) {
    document.getElementById('userLoginEmail').value = found.EM_Email;
    document.getElementById('userLoginCode').value = found.Em_ID;
  }
}

async function handleUserLogin(e) {
  e.preventDefault();
  const email = (document.getElementById('userLoginEmail').value || '').trim().toLowerCase();
  const code = (document.getElementById('userLoginCode').value || '').trim();

  await loadMembersFromSupabase();
  const matched = memberList.find(
    m => (m.EM_Email || '').toLowerCase() === email &&
         (m.Em_ID || '').toUpperCase() === code.toUpperCase()
  );

  if (!matched) {
    showToast('ไม่พบข้อมูลอีเมลหรือรหัสผู้ใช้งานในระบบ กรุณาให้แอดมินเพิ่มผู้ใช้งานในระบบหลังบ้านก่อน', 'error');
    return;
  }

  if (matched.Em_Status === 'inactive') {
    showToast('บัญชีผู้ใช้งานนี้ถูกระงับการใช้งานชั่วคราว', 'error');
    return;
  }

  currentUser = {
    role: 'user',
    name: matched.Em_Name,
    email: matched.EM_Email,
    code: matched.Em_ID
  };
  enterApplication();
  showToast(`ยินดีต้อนรับคุณ ${matched.Em_Name} (${matched.Em_ID}) เข้าสู่ระบบหน้าบ้าน`, 'success');
}

function enterApplication() {
  document.getElementById('authView').style.display = 'none';
  document.getElementById('sessionBar').style.display = 'flex';
  document.getElementById('mainNav').style.display = 'flex';
  document.getElementById('appMain').style.display = 'block';

  document.getElementById('currentUserName').textContent = `${currentUser.name} (${currentUser.code})`;
  const roleBadge = document.getElementById('currentUserRoleBadge');
  roleBadge.textContent = currentUser.role === 'admin' ? 'ADMIN (ระบบหลังบ้าน)' : 'USER (ระบบหน้าบ้าน)';
  roleBadge.className = currentUser.role === 'admin' ? 'role-badge admin' : 'role-badge';
  document.getElementById('currentUserEmailSpan').textContent = currentUser.email ? `✉️ ${currentUser.email}` : '';

  // Role-based access control:
  // "ส่วนผู้ใช้งาน ห้ามเห็นในหน้าเพจที่ 6.รายได้-ค่าใช้จ่าย และ 7 หน้างบกำไรขาดทุน" + หน้าจัดการสมาชิก (page8)
  const adminNavs = document.querySelectorAll('.admin-only-nav');
  adminNavs.forEach(btn => {
    btn.style.display = currentUser.role === 'admin' ? 'inline-flex' : 'none';
  });

  const adminEls = document.querySelectorAll('.admin-only-el');
  adminEls.forEach(el => {
    el.style.display = currentUser.role === 'admin' ? '' : 'none';
  });

  switchPage('page1');
  syncAllData(false);
}

function logoutSystem() {
  currentUser = null;
  document.getElementById('adminUsernameMask').value = '';
  document.getElementById('adminPasswordMask').value = '';
  document.getElementById('adminUserReal').value = '';
  document.getElementById('adminPassReal').value = '';

  document.getElementById('sessionBar').style.display = 'none';
  document.getElementById('mainNav').style.display = 'none';
  document.getElementById('appMain').style.display = 'none';
  document.getElementById('authView').style.display = 'block';
  showToast('ออกจากระบบเรียบร้อยแล้ว', 'info');
}

function switchPage(pageId) {
  // Block non-admin from accessing pages 6, 7, 8
  if (currentUser && currentUser.role !== 'admin' && ['page6', 'page7', 'page8'].includes(pageId)) {
    showToast('ผู้ใช้งานทั่วไปไม่สามารถเข้าถึงหน้านี้ได้ (เฉพาะแอดมินเท่านั้น)', 'error');
    return;
  }

  document.querySelectorAll('.page-section').forEach(sec => sec.classList.remove('active'));
  const target = document.getElementById(pageId);
  if (target) target.classList.add('active');

  document.querySelectorAll('.nav-btn').forEach(btn => {
    btn.classList.toggle('active', btn.getAttribute('data-page') === pageId);
  });
}

// ============================================================================
// IMAGE PLACEHOLDER & BANNER UPLOAD HANDLERS
// ============================================================================
function handleBannerUpload(event) {
  const file = event.target.files && event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    document.getElementById('mainBannerImg').src = e.target.result;
    showToast('เปลี่ยนภาพแบนเนอร์ร้านเครื่องสำอาง AP เรียบร้อยแล้ว', 'success');
  };
  reader.readAsDataURL(file);
}

function previewPageSlotImage(event, previewBoxId) {
  const file = event.target.files && event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    const box = document.getElementById(previewBoxId);
    if (box) {
      box.innerHTML = `<img src="${e.target.result}" alt="Page Custom Image" />`;
      showToast(`แทรกรูปภาพสำเร็จ (${file.name})`, 'success');
    }
  };
  reader.readAsDataURL(file);
}

// ============================================================================
// SUPABASE DATA LOADING & REALTIME SYNC ACROSS ALL DEVICES
// ============================================================================
async function syncAllData(notify = false) {
  try {
    await Promise.all([
      loadInventoryFromSupabase(),
      loadPurchaseSelectFromSupabase(),
      loadPurchaseSummaryFromSupabase(),
      loadSalesSummaryFromSupabase(),
      loadExpensesFromSupabase(),
      loadMembersFromSupabase()
    ]);
    if (notify) {
      showToast('ซิงค์ข้อมูลล่าสุดจากฐานข้อมูล APSHOP1122\'sOrg เรียบร้อยแล้ว', 'success');
    }
  } catch (err) {
    console.error('Sync error:', err);
  }
}

// 1. Load Inventory (Inventorys_TB & Inventory_TB)
async function loadInventoryFromSupabase() {
  const { data, error } = await db
    .from('Inventorys_TB')
    .select('*')
    .order('In_ID', { ascending: true });

  if (error || !data || data.length === 0) {
    const res2 = await db.from('Inventory_TB').select('*').order('In_ID', { ascending: true });
    inventoryList = res2.data || [];
  } else {
    inventoryList = data || [];
  }

  renderInventoryTable();
  renderSalesShelf();
}

// 2. Load Purchase Select:
// "**หน้าซื้อสินค้า จะใช้กับฐานข้อมูลเดิมแต่เรียกใช้คำสั่ง SELECT เพื่อแสดงเฉพาะ In_Bar, In_Menu, In_Price"
async function loadPurchaseSelectFromSupabase() {
  const { data, error } = await db
    .from('Inventorys_TB')
    .select('In_Bar, In_Menu, In_PriceSum, In_Price_Cost')
    .order('In_Bar', { ascending: true });

  if (!error && data) {
    purchaseSelectList = data.map(i => ({
      In_Bar: i.In_Bar,
      In_Menu: i.In_Menu,
      In_Price: getUnitCostPrice(i)
    }));
  } else {
    purchaseSelectList = inventoryList.map(i => ({
      In_Bar: i.In_Bar,
      In_Menu: i.In_Menu,
      In_Price: getUnitCostPrice(i)
    }));
  }
  renderPurchaseShelf();
}

// 3. Load Purchase Summary:
// "**หน้าสรุปยอดซื้อสินค้า จะใช้กับข้อมูลเดิมแต่ใช้คำสั่ง SELECT เพื่อแสดงเฉพาะ In_Date, In_Bar, In_Menu, In_Qty, In_Count_IN, In_Price_Cost, In_Sum_Cost"
async function loadPurchaseSummaryFromSupabase() {
  const { data, error } = await db
    .from('Summary_Purch_TB')
    .select('P_ID, In_Date, In_Bar, In_Menu, In_Qty, In_Count_IN, In_Price_Cost, In_Sum_Cost')
    .order('P_ID', { ascending: false });

  if (!error && data) {
    purchaseSummaryList = data;
  }
  renderPurchaseSummary();
  renderProfitLossPage();
}

// 4. Load Sales Summary (Summary_Self_TB + Inventory_TB analysis):
async function loadSalesSummaryFromSupabase() {
  const { data, error } = await db
    .from('Summary_Self_TB')
    .select('S_ID, S_Date, In_Bar, In_Menu, In_Qty, In_Price_Self, S_Sum, In_Receipt, S_Cashier')
    .order('S_ID', { ascending: false });

  if (!error && data) {
    salesSummaryList = data;
  }
  renderSalesSummary();
  renderProfitLossPage();
}

// 5. Load Income & Expense Summary (Summary_Expans_TB):
async function loadExpensesFromSupabase() {
  const { data, error } = await db
    .from('Summary_Expans_TB')
    .select('*')
    .order('Ex_Date', { ascending: false });

  if (!error && data) {
    expenseSummaryList = data;
  }
  renderExpenseTable();
  renderProfitLossPage();
}

// 6. Load Members (Employee_TB):
async function loadMembersFromSupabase() {
  const { data, error } = await db
    .from('Employee_TB')
    .select('*')
    .order('Em_ID', { ascending: true });

  if (!error && data) {
    memberList = data;
  }
  renderMemberTable();
  populateQuickUserSelect();
  generateNextMemberCode();
}

function populateQuickUserSelect() {
  const sel = document.getElementById('quickUserSelect');
  if (!sel) return;
  const currentVal = sel.value;
  sel.innerHTML = `<option value="">-- เลือกผู้ใช้งานจากฐานข้อมูล Employee_TB (${memberList.length} คน) --</option>`;
  memberList.forEach(m => {
    const opt = document.createElement('option');
    opt.value = m.Em_ID;
    opt.textContent = `${m.Em_Name} | ${m.EM_Email} (${m.Em_ID})`;
    sel.appendChild(opt);
  });
  sel.value = currentVal;
}

// ============================================================================
// PAGE 1: INVENTORY MANAGEMENT (Inventory_TB & Inventorys_TB)
// Core Rule: In_Price_Cost - In_Price_Self = In_Qty
// ============================================================================
function computeRowRemainQty(qtyIn, qtyOut) {
  const mode = document.getElementById('formulaModeSelect') ? document.getElementById('formulaModeSelect').value : 'auto';
  const a = Number(qtyIn || 0);
  const b = Number(qtyOut || 0);
  if (mode === 'custom') {
    const expr = (document.getElementById('customFormulaInput')?.value || 'In_Price_Cost - In_Price_Self')
      .replace(/In_Price_Cost/g, String(a))
      .replace(/In_Price_Self/g, String(b));
    try {
      const evaluated = Function(`"use strict"; return (${expr});`)();
      if (!isNaN(evaluated)) return Number(evaluated);
    } catch (_) {}
  }
  return a - b;
}

function toggleFormulaMode(mode) {
  const input = document.getElementById('customFormulaInput');
  if (!input) return;
  if (mode === 'auto') {
    input.value = 'In_Price_Cost - In_Price_Self';
    input.setAttribute('readonly', 'readonly');
  } else {
    input.removeAttribute('readonly');
    input.focus();
  }
}

async function syncBothInventoryTables(bar, payload) {
  // Keep both Inventory_TB and Inventorys_TB 100% synchronized!
  const qtyIn = Math.max(0, Number(payload.In_Price_Cost ?? 0));
  const qtyOut = Math.max(0, Number(payload.In_Price_Self ?? 0));
  const qtyRemain = computeRowRemainQty(qtyIn, qtyOut);

  const existing = inventoryList.find(x => x.In_Bar === bar) || {};
  const unitCost = payload.In_Price !== undefined ? Number(payload.In_Price) : getUnitCostPrice(existing);
  const unitSell = payload.In_Sell_Price !== undefined ? Number(payload.In_Sell_Price) : getUnitSellPrice(existing);

  const statusText = qtyRemain > 0 ? 'พร้อมขาย' : 'สั่งซื้อ';

  const rowData = {
    In_Count_Cost_IN: Math.round(qtyIn),
    In_Price_Cost: Math.round(qtyIn),
    In_Sum_Cost: Math.round(qtyIn * unitCost),
    In_Count_Self: Math.round(qtyOut),
    In_Price_Self: Math.round(qtyOut),
    In_Sum_Self: Math.round(qtyOut * unitSell),
    In_Qty: Math.round(qtyRemain),
    In_PriceSum: Math.round(unitCost),
    In_SumSelf: Math.round(qtyRemain * unitCost),
    In_Status: statusText,
    In_Date: getTodayStr()
  };
  if (payload.In_Receipt !== undefined) rowData.In_Receipt = payload.In_Receipt;
  if (payload.In_Menu !== undefined) rowData.In_Menu = payload.In_Menu;
  if (payload.In_Unit !== undefined) rowData.In_Unit = payload.In_Unit;

  await Promise.all([
    db.from('Inventory_TB').update(rowData).eq('In_Bar', bar),
    db.from('Inventorys_TB').update(rowData).eq('In_Bar', bar)
  ]);
}

function renderInventoryTable() {
  const tbody = document.getElementById('inventoryTableBody');
  if (!tbody) return;

  const q = (document.getElementById('invSearchInput')?.value || '').trim().toLowerCase();
  const filtered = inventoryList.filter(item => {
    if (!q) return true;
    return (item.In_Bar || '').toLowerCase().includes(q) ||
           (item.In_Menu || '').toLowerCase().includes(q);
  });

  let totalIn = 0;
  let totalOut = 0;
  let totalRemain = 0;

  inventoryList.forEach(item => {
    const qIn = Number(item.In_Price_Cost || 0);
    const qOut = Number(item.In_Price_Self || 0);
    const qRem = qIn - qOut;
    totalIn += qIn;
    totalOut += qOut;
    totalRemain += qRem;
  });

  document.getElementById('invStatTotalItems').textContent = `${inventoryList.length} รายการ`;
  document.getElementById('invStatTotalIn').textContent = `${totalIn.toLocaleString()} ชิ้น`;
  document.getElementById('invStatTotalOut').textContent = `${totalOut.toLocaleString()} ชิ้น`;
  document.getElementById('invStatTotalRemain').textContent = `${totalRemain.toLocaleString()} ชิ้น`;

  if (filtered.length === 0) {
    tbody.innerHTML = `<tr><td colspan="16" style="text-align:center; color:var(--color-muted);">ไม่พบรายการสินค้า</td></tr>`;
    return;
  }

  const isAdmin = currentUser && currentUser.role === 'admin';

  tbody.innerHTML = filtered.map(item => {
    const unitCost = getUnitCostPrice(item);
    const unitSell = getUnitSellPrice(item);
    const qtyIn = Number(item.In_Price_Cost || 0);
    const countCostIn = item.In_Count_Cost_IN !== null && item.In_Count_Cost_IN !== undefined ? Number(item.In_Count_Cost_IN) : qtyIn;
    const qtyOut = Number(item.In_Price_Self || 0);
    const countSelf = item.In_Count_Self !== null && item.In_Count_Self !== undefined ? Number(item.In_Count_Self) : qtyOut;
    const qtyRemain = qtyIn - qtyOut;
    const sumCost = Number(item.In_Sum_Cost || (qtyIn * unitCost));
    const sumSell = Number(item.In_Sum_Self || (qtyOut * unitSell));
    const priceSum = Number(item.In_PriceSum ?? unitCost);
    const sumSelf = Number(item.In_SumSelf ?? (qtyRemain * priceSum));
    const receiptText = item.In_Receipt || '-';
    const statusLabel = qtyRemain > 0 ? 'พร้อมขาย' : (item.In_Status || 'สั่งซื้อ');
    const badgeClass = qtyRemain > 0 ? 'badge-ready' : 'badge-order';

    return `
      <tr>
        <td><strong>${item.In_ID ?? '-'}</strong></td>
        <td><span class="shelf-bar-tag">${item.In_Bar}</span></td>
        <td><strong>${item.In_Menu}</strong></td>
        <td>${item.In_Unit || 'ชิ้น'}</td>
        <td>${countCostIn}</td>
        <td style="background:#FFF5F5; font-weight:700; color:var(--color-secondary);">
          ${qtyIn}
        </td>
        <td>${formatMoney(sumCost)}</td>
        <td>${countSelf}</td>
        <td style="background:#FFF5F5; font-weight:700; color:var(--color-primary);">
          ${qtyOut}
        </td>
        <td>${formatMoney(sumSell)}</td>
        <td style="background:#FEF2F2; font-weight:700; font-size:15px; color:${qtyRemain < 0 ? 'var(--color-danger)' : 'var(--color-dark)'};">
          ${qtyRemain}
        </td>
        <td>${formatMoney(priceSum)}</td>
        <td>${formatMoney(sumSelf)}</td>
        <td>${receiptText}</td>
        <td><span class="badge ${badgeClass}">${statusLabel}</span></td>
        <td>
          <div style="display:flex; gap:6px;">
            ${isAdmin ? `<button class="btn btn-sm btn-outline" onclick="openEditInventoryModal('${item.In_Bar}')">✏️ แก้ไข</button>` : ''}
            <button class="btn btn-sm btn-danger" onclick="deleteInventoryItem('${item.In_Bar}', '${item.In_Menu}')">🗑️ ลบ</button>
          </div>
        </td>
      </tr>
    `;
  }).join('');
}

async function resetAllInventoryZero() {
  for (const item of inventoryList) {
    await syncBothInventoryTables(item.In_Bar, {
      In_Price_Cost: 0,
      In_Price_Self: 0,
      In_Qty: 0
    });
  }
  await loadInventoryFromSupabase();
  showToast('กำหนดยอดรับเข้า (In_Price_Cost), จ่ายออก (In_Price_Self) และคงเหลือ (In_Qty) เป็น 0 เรียบร้อยแล้ว', 'success');
}

async function recalculateAllInventoryFormulas() {
  for (const item of inventoryList) {
    await syncBothInventoryTables(item.In_Bar, {
      In_Price_Cost: Number(item.In_Price_Cost || 0),
      In_Price_Self: Number(item.In_Price_Self || 0)
    });
  }
  await loadInventoryFromSupabase();
  showToast('คำนวณสูตร In_Price_Cost - In_Price_Self = In_Qty และอัปเดตฐานข้อมูลเรียบร้อยแล้ว', 'success');
}

function openInventoryModal() {
  document.getElementById('invModalMode').value = 'create';
  document.getElementById('invModalTitle').textContent = '➕ เพิ่มรายการสินค้าใหม่ (อ้างอิงโครงสร้างไฟล์ Excel)';
  const nextId = inventoryList.length > 0 ? Math.max(...inventoryList.map(i => Number(i.In_ID || 0))) + 1 : 1;
  document.getElementById('invModalId').value = nextId;
  document.getElementById('invModalBar').value = `PA${String(nextId).padStart(3, '0')}`;
  document.getElementById('invModalBar').removeAttribute('readonly');
  document.getElementById('invModalMenu').value = '';
  document.getElementById('invModalUnit').value = 'ชิ้น';
  document.getElementById('invModalCostPrice').value = '';
  document.getElementById('invModalSellPrice').value = '';
  document.getElementById('invModalQtyIn').value = 0;
  document.getElementById('invModalQtyOut').value = 0;
  document.getElementById('invModalQtyRemain').value = 0;
  document.getElementById('inventoryModal').classList.add('open');
}

function openEditInventoryModal(bar) {
  const item = inventoryList.find(x => x.In_Bar === bar);
  if (!item) return;
  document.getElementById('invModalMode').value = 'edit';
  document.getElementById('invModalTitle').textContent = `✏️ แก้ไขรายการสินค้า (${item.In_Bar})`;
  document.getElementById('invModalId').value = item.In_ID;
  document.getElementById('invModalBar').value = item.In_Bar;
  document.getElementById('invModalBar').setAttribute('readonly', 'readonly');
  document.getElementById('invModalMenu').value = item.In_Menu;
  document.getElementById('invModalUnit').value = item.In_Unit || 'ชิ้น';
  document.getElementById('invModalCostPrice').value = getUnitCostPrice(item);
  document.getElementById('invModalSellPrice').value = getUnitSellPrice(item);
  document.getElementById('invModalQtyIn').value = Number(item.In_Price_Cost || 0);
  document.getElementById('invModalQtyOut').value = Number(item.In_Price_Self || 0);
  updateModalRemainPreview();
  document.getElementById('inventoryModal').classList.add('open');
}

function closeInventoryModal() {
  document.getElementById('inventoryModal').classList.remove('open');
}

function updateModalRemainPreview() {
  const qIn = Number(document.getElementById('invModalQtyIn').value || 0);
  const qOut = Number(document.getElementById('invModalQtyOut').value || 0);
  document.getElementById('invModalQtyRemain').value = qIn - qOut;
}

async function handleSaveInventoryItem(e) {
  e.preventDefault();
  const mode = document.getElementById('invModalMode').value;
  const id = Number(document.getElementById('invModalId').value);
  const bar = document.getElementById('invModalBar').value.trim().toUpperCase();
  const menu = document.getElementById('invModalMenu').value.trim();
  const unit = document.getElementById('invModalUnit').value.trim();
  const costPrice = Number(document.getElementById('invModalCostPrice').value || 0);
  const sellPrice = Number(document.getElementById('invModalSellPrice').value || 0);
  const qtyIn = Number(document.getElementById('invModalQtyIn').value || 0);
  const qtyOut = Number(document.getElementById('invModalQtyOut').value || 0);
  const qtyRemain = qtyIn - qtyOut;

  EXCEL_MAP[bar] = { In_ID: id, In_Bar: bar, In_Menu: menu, In_Unit: unit, cost: costPrice, sell: sellPrice };

  const rowPayload = {
    In_ID: id,
    In_Bar: bar,
    In_Menu: menu,
    In_Unit: unit,
    In_Count_Cost_IN: Math.round(qtyIn),
    In_Price_Cost: Math.round(qtyIn),
    In_Sum_Cost: Math.round(qtyIn * costPrice),
    In_Count_Self: Math.round(qtyOut),
    In_Price_Self: Math.round(qtyOut),
    In_Sum_Self: Math.round(qtyOut * sellPrice),
    In_Qty: Math.round(qtyRemain),
    In_PriceSum: Math.round(costPrice),
    In_SumSelf: Math.round(qtyRemain * costPrice),
    In_Status: qtyRemain > 0 ? 'พร้อมขาย' : 'สั่งซื้อ',
    In_Date: getTodayStr()
  };

  if (mode === 'create') {
    const { error } = await db.from('Inventory_TB').insert([rowPayload]);
    await db.from('Inventorys_TB').insert([rowPayload]);
    if (error) {
      showToast(`เกิดข้อผิดพลาดในการเพิ่มสินค้า: ${error.message}`, 'error');
      return;
    }
    showToast(`เพิ่มสินค้า ${bar} (${menu}) สำเร็จ!`, 'success');
  } else {
    await syncBothInventoryTables(bar, rowPayload);
    showToast(`อัปเดตข้อมูลสินค้า ${bar} (${menu}) สำเร็จ!`, 'success');
  }

  closeInventoryModal();
  await syncAllData(false);
}

async function deleteInventoryItem(bar, menuName) {
  await Promise.all([
    db.from('Inventory_TB').delete().eq('In_Bar', bar),
    db.from('Inventorys_TB').delete().eq('In_Bar', bar)
  ]);
  await syncAllData(false);
  showToast(`ลบสินค้า ${bar} (${menuName}) ออกจากคลังสินค้าเรียบร้อยแล้ว`, 'success');
}

// ============================================================================
// PAGE 2: PURCHASE POS SYSTEM (ระบบซื้อสินค้าแบบร้านสะดวกซื้อ)
// Uses SELECT In_Bar, In_Menu, In_Price
// Rule #2: หน้ารายการซื้อ เมื่อเพิ่มสินค้าเข้าไปช่อง จำนวน(Q'ty) ต้องเชื่อมกับหน้าคลังสินค้าในคอลัมน์ In_Price_Cost
// ============================================================================
function renderPurchaseShelf() {
  const grid = document.getElementById('purchaseShelfGrid');
  if (!grid) return;
  const filter = (document.getElementById('purchShelfFilter')?.value || '').trim().toLowerCase();

  // Merge purchaseSelectList (In_Bar, In_Menu, In_Price) with stock info
  const source = purchaseSelectList.length > 0 ? purchaseSelectList : inventoryList;
  const filtered = source.filter(p => {
    if (!filter) return true;
    return (p.In_Bar || '').toLowerCase().includes(filter) ||
           (p.In_Menu || '').toLowerCase().includes(filter);
  });

  grid.innerHTML = filtered.map(p => {
    const invItem = inventoryList.find(x => x.In_Bar === p.In_Bar) || {};
    const costPrice = p.In_Price !== undefined && Number(p.In_Price) > 0 ? Number(p.In_Price) : getUnitCostPrice(invItem);
    const curCostQty = Number(invItem.In_Price_Cost || 0);
    const curRemain = Number(invItem.In_Qty || 0);

    return `
      <div class="shelf-item-card" onclick="addProductToPurchaseCart('${p.In_Bar}', 1)">
        <div>
          <span class="shelf-bar-tag">${p.In_Bar}</span>
          <div class="shelf-name">${p.In_Menu}</div>
        </div>
        <div>
          <div class="shelf-price">ราคาทุน ฿${formatMoney(costPrice)}</div>
          <div class="shelf-stock">รับเข้าสะสม (In_Price_Cost): <strong>${curCostQty}</strong> | คงเหลือ: <strong>${curRemain}</strong></div>
        </div>
      </div>
    `;
  }).join('');
}

function handlePurchaseBarcodeScan(e) {
  e.preventDefault();
  const barInput = document.getElementById('purchBarcodeInput');
  const qtyInput = document.getElementById('purchScanQty');
  const bar = (barInput.value || '').trim().toUpperCase();
  const qty = Math.max(1, Number(qtyInput.value || 1));

  if (!bar) return;
  const added = addProductToPurchaseCart(bar, qty);
  if (added) {
    barInput.value = '';
    qtyInput.value = '1';
    barInput.focus();
  }
}

function addProductToPurchaseCart(bar, qtyToAdd = 1) {
  const pSelect = purchaseSelectList.find(x => x.In_Bar.toUpperCase() === bar.toUpperCase());
  const invItem = inventoryList.find(x => x.In_Bar.toUpperCase() === bar.toUpperCase());
  const item = pSelect || invItem;

  if (!item) {
    showToast(`ไม่พบรหัสบาร์โค้ด ${bar} ในระบบคลังสินค้า`, 'error');
    return false;
  }

  const unitCost = pSelect && Number(pSelect.In_Price) > 0 ? Number(pSelect.In_Price) : getUnitCostPrice(invItem || item);
  const existing = purchaseCart.find(c => c.In_Bar === item.In_Bar);

  if (existing) {
    existing.qty += Number(qtyToAdd);
    existing.sumCost = existing.qty * existing.unitCost;
  } else {
    purchaseCart.push({
      In_Bar: item.In_Bar,
      In_Menu: item.In_Menu,
      In_Unit: (invItem && invItem.In_Unit) || 'ชิ้น',
      unitCost: unitCost,
      qty: Number(qtyToAdd),
      sumCost: Number(qtyToAdd) * unitCost
    });
  }

  renderPurchaseCart();
  return true;
}

function updatePurchaseCartQty(bar, newQty) {
  const q = Number(newQty);
  if (q <= 0) {
    removeFromPurchaseCart(bar);
    return;
  }
  const row = purchaseCart.find(c => c.In_Bar === bar);
  if (row) {
    row.qty = q;
    row.sumCost = row.qty * row.unitCost;
    renderPurchaseCart();
  }
}

function removeFromPurchaseCart(bar) {
  purchaseCart = purchaseCart.filter(c => c.In_Bar !== bar);
  renderPurchaseCart();
}

function clearPurchaseCart() {
  purchaseCart = [];
  renderPurchaseCart();
}

function renderPurchaseCart() {
  const tbody = document.getElementById('purchCartBody');
  if (!tbody) return;

  let totalQty = 0;
  let totalCost = 0;

  if (purchaseCart.length === 0) {
    tbody.innerHTML = `<tr><td colspan="6" style="text-align:center; color:var(--color-muted);">ยังไม่มีรายการในตะกร้าซื้อสินค้า</td></tr>`;
  } else {
    tbody.innerHTML = purchaseCart.map(c => {
      totalQty += c.qty;
      totalCost += c.sumCost;
      return `
        <tr>
          <td><strong>${c.In_Bar}</strong></td>
          <td>${c.In_Menu}</td>
          <td>฿${formatMoney(c.unitCost)}</td>
          <td>
            <input type="number" class="form-control" style="width:76px; padding:4px 8px;" min="1" value="${c.qty}"
              onchange="updatePurchaseCartQty('${c.In_Bar}', this.value)" />
          </td>
          <td><strong>฿${formatMoney(c.sumCost)}</strong></td>
          <td>
            <button class="btn btn-sm btn-danger" onclick="removeFromPurchaseCart('${c.In_Bar}')">✕</button>
          </td>
        </tr>
      `;
    }).join('');
  }

  document.getElementById('purchNetTotalDisplay').textContent = `฿${formatMoney(totalCost)}`;
  document.getElementById('purchTotalQtySpan').textContent = `${totalQty} ชิ้น (เชื่อมเข้า In_Price_Cost)`;
  document.getElementById('purchTotalAmountSpan').textContent = `${formatMoney(totalCost)} บาท`;
}

async function submitPurchaseOrder() {
  if (purchaseCart.length === 0) {
    showToast('กรุณาสแกนหรือเลือกสินค้าเข้าตะกร้าซื้อสินค้าก่อนบันทึก', 'error');
    return;
  }

  const purchDate = document.getElementById('purchDateInput').value || getTodayStr();
  const nowTime = new Date().toTimeString().slice(0, 8);
  const receiptNo = `PO-${Date.now().toString().slice(-6)}`;
  document.getElementById('purchReceiptNoDisplay').textContent = `เลขที่ใบรับ: ${receiptNo}`;

  const buyerName = currentUser ? currentUser.name : 'พนักงานคลังสินค้า';

  // 1. Insert rows into Summary_Purch_TB
  const rowsToInsert = purchaseCart.map(c => ({
    P_Date: purchDate,
    P_Time: nowTime,
    In_Date: purchDate,
    In_Receipt: receiptNo,
    In_Bar: c.In_Bar,
    In_Menu: c.In_Menu,
    In_Unit: c.In_Unit,
    In_Qty: c.qty,
    In_Count_IN: c.qty,
    In_Price_Cost: c.unitCost,
    In_Sum_Cost: c.sumCost,
    P_Sum: c.sumCost,
    P_Buyer: buyerName
  }));

  const { error } = await db.from('Summary_Purch_TB').insert(rowsToInsert);
  if (error) {
    showToast(`เกิดข้อผิดพลาดในการบันทึกยอดซื้อ: ${error.message}`, 'error');
    return;
  }

  // 2. Rule #2: "หน้ารายการซื้อ เมื่อเพิ่มสินเข้าไปช่อง จำนวน(Q'ty) ต้อง เชื่อมกับ หน้า คลังสินค้า ใน คอลัมน์ In_Price_Cost"
  // And Rule #1: "In_Price_Cost - In_Price_Self = In_Qty"
  for (const c of purchaseCart) {
    const inv = inventoryList.find(x => x.In_Bar === c.In_Bar);
    const curCostQty = Number(inv?.In_Price_Cost || 0);
    const curSellQty = Number(inv?.In_Price_Self || 0);
    const newCostQty = curCostQty + Number(c.qty);
    await syncBothInventoryTables(c.In_Bar, {
      In_Price_Cost: newCostQty,
      In_Price_Self: curSellQty,
      In_Receipt: receiptNo
    });
  }

  purchaseCart = [];
  renderPurchaseCart();
  await syncAllData(false);
  await syncDailyIncomeExpenseToSupabase(false);
  showToast(`บันทึกรับสินค้าเข้าคลัง (${receiptNo}) สำเร็จ! อัปเดตคอลัมน์ In_Price_Cost และคำนวณ In_Qty เรียบร้อยแล้ว`, 'success');
}

// ============================================================================
// PAGE 3: PURCHASE SUMMARY (หน้าสรุปยอดซื้อสินค้า — Summary_Purch_TB)
// SELECT In_Date, In_Bar, In_Menu, In_Qty, In_Count_IN, In_Price_Cost, In_Sum_Cost
// Rule #3: เมื่อมีการลบสินค้า ต้องหักจำนวน (Q'ty) ในหน้าคลังสินค้า ในคอลัมน์ In_Price_Cost
// ============================================================================
function renderPurchaseSummary() {
  const tbody = document.getElementById('purchSummaryTableBody');
  if (!tbody) return;

  const filterDate = document.getElementById('filterPurchDate')?.value || '';
  const filtered = purchaseSummaryList.filter(row => {
    if (!filterDate) return true;
    return (row.In_Date || '') === filterDate;
  });

  let sumQty = 0;
  let sumCountIn = 0;
  let sumCost = 0;

  if (filtered.length === 0) {
    tbody.innerHTML = `<tr><td colspan="8" style="text-align:center; color:var(--color-muted);">ยังไม่มีรายการสรุปยอดซื้อสินค้าในวันที่เลือก</td></tr>`;
  } else {
    tbody.innerHTML = filtered.map(row => {
      const q = Number(row.In_Qty || 0);
      const cIn = Number(row.In_Count_IN || q);
      const pCost = Number(row.In_Price_Cost || 0);
      const sCost = Number(row.In_Sum_Cost || q * pCost);

      sumQty += q;
      sumCountIn += cIn;
      sumCost += sCost;

      return `
        <tr>
          <td>${row.In_Date || '-'}</td>
          <td><span class="shelf-bar-tag">${row.In_Bar}</span></td>
          <td><strong>${row.In_Menu}</strong></td>
          <td><strong>${q}</strong></td>
          <td>${cIn}</td>
          <td>฿${formatMoney(pCost)}</td>
          <td style="color:var(--color-primary); font-weight:700;">฿${formatMoney(sCost)}</td>
          <td>
            <button class="btn btn-sm btn-danger" onclick="deletePurchaseSummaryRow(${row.P_ID}, '${row.In_Bar}', ${q}, '${row.In_Menu}')">
              🗑️ ลบรายการ (หัก In_Price_Cost -${q})
            </button>
          </td>
        </tr>
      `;
    }).join('');
  }

  document.getElementById('sumPurchCount').textContent = `${filtered.length} รายการ`;
  document.getElementById('sumPurchQty').textContent = `${sumQty.toLocaleString()} ชิ้น`;
  document.getElementById('sumPurchTotalCost').textContent = `฿${formatMoney(sumCost)}`;

  document.getElementById('footPurchQty').textContent = sumQty.toLocaleString();
  document.getElementById('footPurchCountIn').textContent = sumCountIn.toLocaleString();
  document.getElementById('footPurchSumCost').textContent = `฿${formatMoney(sumCost)}`;
}

async function deletePurchaseSummaryRow(pId, bar, qtyToDeduct, menuName) {
  const { error } = await db.from('Summary_Purch_TB').delete().eq('P_ID', pId);
  if (error) {
    showToast(`ลบรายการซื้อไม่สำเร็จ: ${error.message}`, 'error');
    return;
  }

  // Deduct Q'ty from In_Price_Cost in Inventory_TB & Inventorys_TB
  const inv = inventoryList.find(x => x.In_Bar === bar);
  if (inv) {
    const curCostQty = Number(inv.In_Price_Cost || 0);
    const curSellQty = Number(inv.In_Price_Self || 0);
    const updatedCostQty = Math.max(0, curCostQty - Number(qtyToDeduct || 0));
    await syncBothInventoryTables(bar, {
      In_Price_Cost: updatedCostQty,
      In_Price_Self: curSellQty
    });
  }

  await syncAllData(false);
  await syncDailyIncomeExpenseToSupabase(false);
  showToast(`ลบรายการซื้อ ${menuName} และหักจำนวน ${qtyToDeduct} ชิ้นออกจาก In_Price_Cost ในคลังสินค้าเรียบร้อยแล้ว`, 'success');
}

// ============================================================================
// PAGE 4: SALES POS SYSTEM (หน้าขายสินค้า POS แบบร้านสะดวกซื้อ + คำนวณเงินทอน)
// Rule #4: หน้าขายสินค้า(POS) เมื่อเพิ่มสินค้าเข้าไปช่อง จำนวน ต้องเชื่อมกับหน้าคลังสินค้าในคอลัมน์ In_Price_Self
// ============================================================================
function renderSalesShelf() {
  const grid = document.getElementById('salesShelfGrid');
  if (!grid) return;
  const filter = (document.getElementById('salesShelfFilter')?.value || '').trim().toLowerCase();

  const filtered = inventoryList.filter(item => {
    if (!filter) return true;
    return (item.In_Bar || '').toLowerCase().includes(filter) ||
           (item.In_Menu || '').toLowerCase().includes(filter);
  });

  grid.innerHTML = filtered.map(item => {
    const sellPrice = getUnitSellPrice(item);
    const qtyIn = Number(item.In_Price_Cost || 0);
    const qtyOut = Number(item.In_Price_Self || 0);
    const qtyRemain = qtyIn - qtyOut;

    return `
      <div class="shelf-item-card" onclick="addProductToSalesCart('${item.In_Bar}', 1)">
        <div>
          <span class="shelf-bar-tag">${item.In_Bar}</span>
          <div class="shelf-name">${item.In_Menu}</div>
        </div>
        <div>
          <div class="shelf-price">ราคาขาย ฿${formatMoney(sellPrice)}</div>
          <div class="shelf-stock">
            ขายแล้ว (In_Price_Self): <strong>${qtyOut}</strong> | คงเหลือ: <strong style="color:${qtyRemain > 0 ? 'var(--color-success)' : 'var(--color-danger)'};">${qtyRemain}</strong>
          </div>
        </div>
      </div>
    `;
  }).join('');
}

function handleSalesBarcodeScan(e) {
  e.preventDefault();
  const barInput = document.getElementById('salesBarcodeInput');
  const qtyInput = document.getElementById('salesScanQty');
  const bar = (barInput.value || '').trim().toUpperCase();
  const qty = Math.max(1, Number(qtyInput.value || 1));

  if (!bar) return;
  const added = addProductToSalesCart(bar, qty);
  if (added) {
    barInput.value = '';
    qtyInput.value = '1';
    barInput.focus();
  }
}

function addProductToSalesCart(bar, qtyToAdd = 1) {
  const item = inventoryList.find(x => x.In_Bar.toUpperCase() === bar.toUpperCase());
  if (!item) {
    showToast(`ไม่พบรหัสบาร์โค้ด ${bar} ในระบบ`, 'error');
    return false;
  }

  const unitSell = getUnitSellPrice(item);
  const existing = salesCart.find(c => c.In_Bar === item.In_Bar);

  if (existing) {
    existing.qty += Number(qtyToAdd);
    existing.sumSell = existing.qty * existing.unitSell;
  } else {
    salesCart.push({
      In_Bar: item.In_Bar,
      In_Menu: item.In_Menu,
      In_Unit: item.In_Unit || 'ชิ้น',
      unitSell: unitSell,
      qty: Number(qtyToAdd),
      sumSell: Number(qtyToAdd) * unitSell
    });
  }

  renderSalesCart();
  return true;
}

function updateSalesCartQty(bar, newQty) {
  const q = Number(newQty);
  if (q <= 0) {
    removeFromSalesCart(bar);
    return;
  }
  const row = salesCart.find(c => c.In_Bar === bar);
  if (row) {
    row.qty = q;
    row.sumSell = row.qty * row.unitSell;
    renderSalesCart();
  }
}

function removeFromSalesCart(bar) {
  salesCart = salesCart.filter(c => c.In_Bar !== bar);
  renderSalesCart();
}

function clearSalesCart() {
  salesCart = [];
  document.getElementById('cashReceivedInput').value = '';
  renderSalesCart();
}

function getSalesCartTotalAmount() {
  return salesCart.reduce((acc, c) => acc + c.sumSell, 0);
}

function renderSalesCart() {
  const tbody = document.getElementById('salesCartBody');
  if (!tbody) return;

  let totalQty = 0;
  let totalAmount = 0;

  if (salesCart.length === 0) {
    tbody.innerHTML = `<tr><td colspan="6" style="text-align:center; color:var(--color-muted);">สแกนบาร์โค้ดหรือคลิกสินค้าเพื่อเริ่มขาย</td></tr>`;
  } else {
    tbody.innerHTML = salesCart.map(c => {
      totalQty += c.qty;
      totalAmount += c.sumSell;
      return `
        <tr>
          <td><strong>${c.In_Bar}</strong></td>
          <td>${c.In_Menu}</td>
          <td>฿${formatMoney(c.unitSell)}</td>
          <td>
            <input type="number" class="form-control" style="width:76px; padding:4px 8px;" min="1" value="${c.qty}"
              onchange="updateSalesCartQty('${c.In_Bar}', this.value)" />
          </td>
          <td><strong>฿${formatMoney(c.sumSell)}</strong></td>
          <td>
            <button class="btn btn-sm btn-danger" onclick="removeFromSalesCart('${c.In_Bar}')">✕</button>
          </td>
        </tr>
      `;
    }).join('');
  }

  document.getElementById('salesNetTotalDisplay').textContent = `฿${formatMoney(totalAmount)}`;
  document.getElementById('salesTotalQtySpan').textContent = `${totalQty} ชิ้น (เชื่อมเข้า In_Price_Self)`;
  document.getElementById('salesTotalAmountSpan').textContent = `${formatMoney(totalAmount)} บาท`;
  calculatePosChange();
}

function setQuickCash(val) {
  const total = getSalesCartTotalAmount();
  const input = document.getElementById('cashReceivedInput');
  if (val === 'exact') {
    input.value = total;
  } else {
    input.value = Number(input.value || 0) + Number(val);
  }
  calculatePosChange();
}

function calculatePosChange() {
  const total = getSalesCartTotalAmount();
  const cashReceived = Number(document.getElementById('cashReceivedInput')?.value || 0);
  const change = cashReceived - total;

  const changeSpan = document.getElementById('salesChangeAmountSpan');
  const ledChange = document.getElementById('salesChangeLedDisplay');

  if (total === 0 && cashReceived === 0) {
    changeSpan.textContent = '0.00 บาท';
    changeSpan.style.color = 'var(--color-success)';
    ledChange.textContent = '฿0.00';
    return change;
  }

  if (change < 0) {
    changeSpan.textContent = `ยังขาดอีก ${formatMoney(Math.abs(change))} บาท`;
    changeSpan.style.color = 'var(--color-danger)';
    ledChange.textContent = `-฿${formatMoney(Math.abs(change))}`;
  } else {
    changeSpan.textContent = `${formatMoney(change)} บาท`;
    changeSpan.style.color = 'var(--color-success)';
    ledChange.textContent = `฿${formatMoney(change)}`;
  }
  return change;
}

async function submitSalesOrder() {
  if (salesCart.length === 0) {
    showToast('กรุณาสแกนหรือเลือกสินค้าเข้าบิลขายก่อนชำระเงิน', 'error');
    return;
  }

  const totalAmount = getSalesCartTotalAmount();
  const cashInput = document.getElementById('cashReceivedInput');
  let cashReceived = Number(cashInput.value || 0);

  // If user didn't type cash yet, default to exact amount or validate if partial was typed
  if (cashInput.value === '') {
    cashReceived = totalAmount;
    cashInput.value = totalAmount;
  }

  if (cashReceived < totalAmount) {
    showToast(`รับเงินมาไม่พอชำระ! ยอดสุทธิ ${formatMoney(totalAmount)} บาท แต่รับมา ${formatMoney(cashReceived)} บาท`, 'error');
    return;
  }

  const changeAmount = cashReceived - totalAmount;
  const salesDate = document.getElementById('salesDateInput').value || getTodayStr();
  const receiptNo = `INV-${Date.now().toString().slice(-6)}`;
  document.getElementById('salesReceiptNoDisplay').textContent = `ใบเสร็จ: ${receiptNo}`;

  const cashierName = currentUser ? currentUser.name : 'แคชเชียร์หน้าร้าน';

  // 1. Insert into Summary_Self_TB
  const rowsToInsert = salesCart.map(c => ({
    S_Date: salesDate,
    In_Bar: c.In_Bar,
    In_Menu: c.In_Menu,
    In_Qty: String(c.qty),
    In_Price_Self: c.unitSell,
    S_Sum: c.sumSell,
    In_Receipt: receiptNo,
    S_Cashier: cashierName
  }));

  const { error } = await db.from('Summary_Self_TB').insert(rowsToInsert);
  if (error) {
    showToast(`บันทึกการขายไม่สำเร็จ: ${error.message}`, 'error');
    return;
  }

  // 2. Rule #4: "หน้าขายสินค้า(POS) เมื่อเพิ่มสินเข้าไปช่อง จำนวน ต้อง เชื่อมกับ หน้า คลังสินค้า ใน คอลัมน์ In_Price_Self"
  // And Rule #1: "In_Price_Cost - In_Price_Self = In_Qty"
  for (const c of salesCart) {
    const inv = inventoryList.find(x => x.In_Bar === c.In_Bar);
    const curCostQty = Number(inv?.In_Price_Cost || 0);
    const curSellQty = Number(inv?.In_Price_Self || 0);
    const newSellQty = curSellQty + Number(c.qty);
    await syncBothInventoryTables(c.In_Bar, {
      In_Price_Cost: curCostQty,
      In_Price_Self: newSellQty,
      In_Receipt: receiptNo
    });
  }

  salesCart = [];
  cashInput.value = '';
  renderSalesCart();
  await syncAllData(false);
  await syncDailyIncomeExpenseToSupabase(false);
  showToast(
    `ชำระเงินสำเร็จ (${receiptNo})! ยอดสุทธิ ฿${formatMoney(totalAmount)} | รับเงิน ฿${formatMoney(cashReceived)} | เงินทอน ฿${formatMoney(changeAmount)} (อัปเดต In_Price_Self และ In_Qty แล้ว)`,
    'success'
  );
}

// ============================================================================
// PAGE 5: SALES SUMMARY (หน้าสรุปยอดขายสินค้า — Summary_Self_TB)
// Rule #5 (and #3): เมื่อลบสินค้าออกไปแล้ว ให้หักจำนวนในหน้าคลังสินค้าในคอลัมน์ In_Price_Self (หรือ In_Price_Cost)
// ============================================================================
function renderSalesSummary() {
  const tbody = document.getElementById('salesSummaryTableBody');
  if (!tbody) return;

  const filterDate = document.getElementById('filterSalesDate')?.value || '';
  const filtered = salesSummaryList.filter(row => {
    if (!filterDate) return true;
    return (row.S_Date || '') === filterDate;
  });

  let totalQty = 0;
  let totalSum = 0;

  if (filtered.length === 0) {
    tbody.innerHTML = `<tr><td colspan="10" style="text-align:center; color:var(--color-muted);">ยังไม่มีรายการสรุปยอดขายสินค้าในวันที่เลือก</td></tr>`;
  } else {
    tbody.innerHTML = filtered.map(row => {
      const q = Number(row.In_Qty || 0);
      const pSelf = Number(row.In_Price_Self || 0);
      const sSum = Number(row.S_Sum || q * pSelf);

      totalQty += q;
      totalSum += sSum;

      return `
        <tr>
          <td><strong>${row.S_ID}</strong></td>
          <td>${row.S_Date || '-'}</td>
          <td><span class="shelf-bar-tag">${row.In_Bar}</span></td>
          <td><strong>${row.In_Menu}</strong></td>
          <td><strong>${q}</strong></td>
          <td>฿${formatMoney(pSelf)}</td>
          <td style="color:var(--color-success); font-weight:700;">฿${formatMoney(sSum)}</td>
          <td><code>${row.In_Receipt || '-'}</code></td>
          <td>${row.S_Cashier || '-'}</td>
          <td>
            <button class="btn btn-sm btn-danger" onclick="deleteSalesSummaryRow(${row.S_ID}, '${row.In_Bar}', ${q}, '${row.In_Menu}')">
              🗑️ ลบรายการขาย (-${q})
            </button>
          </td>
        </tr>
      `;
    }).join('');
  }

  document.getElementById('sumSalesCount').textContent = `${filtered.length} รายการ`;
  document.getElementById('sumSalesQty').textContent = `${totalQty.toLocaleString()} ชิ้น`;
  document.getElementById('sumSalesTotalAmount').textContent = `฿${formatMoney(totalSum)}`;

  document.getElementById('footSalesQty').textContent = totalQty.toLocaleString();
  document.getElementById('footSalesSum').textContent = `฿${formatMoney(totalSum)}`;
}

async function deleteSalesSummaryRow(sId, bar, qtyToDeduct, menuName) {
  const { error } = await db.from('Summary_Self_TB').delete().eq('S_ID', sId);
  if (error) {
    showToast(`ลบรายการขายไม่สำเร็จ: ${error.message}`, 'error');
    return;
  }

  const targetCol = document.getElementById('salesDeleteDeductTarget')?.value || 'In_Price_Self';
  const inv = inventoryList.find(x => x.In_Bar === bar);
  if (inv) {
    let curCostQty = Number(inv.In_Price_Cost || 0);
    let curSellQty = Number(inv.In_Price_Self || 0);
    const d = Number(qtyToDeduct || 0);

    if (targetCol === 'In_Price_Self') {
      curSellQty = Math.max(0, curSellQty - d);
    } else if (targetCol === 'In_Price_Cost') {
      curCostQty = Math.max(0, curCostQty - d);
    } else if (targetCol === 'both') {
      curSellQty = Math.max(0, curSellQty - d);
      curCostQty = Math.max(0, curCostQty - d);
    }

    await syncBothInventoryTables(bar, {
      In_Price_Cost: curCostQty,
      In_Price_Self: curSellQty
    });
  }

  await syncAllData(false);
  await syncDailyIncomeExpenseToSupabase(false);
  showToast(`ลบรายการขาย ${menuName} และหักจำนวน ${qtyToDeduct} ชิ้นในคลังสินค้า (${targetCol}) เรียบร้อยแล้ว`, 'success');
}

// ============================================================================
// PAGE 6: DAILY INCOME & EXPENSE SUMMARY (Summary_Expans_TB) — ADMIN ONLY
// Rule #6: หน้ารายได้-ค่าใช้จ่าย ช่องค่าใช้จ่ายอื่นๆ เปลี่ยนเป็นการกรอกข้อมูลเอง และสร้างแบบฟอร์มกรอกข้อมูลให้
// ============================================================================
function computeDailyTotalsByDate(dateStr) {
  const daySales = salesSummaryList
    .filter(s => (s.S_Date || '') === dateStr)
    .reduce((acc, s) => acc + Number(s.S_Sum || 0), 0);

  const dayPurch = purchaseSummaryList
    .filter(p => (p.In_Date || '') === dateStr)
    .reduce((acc, p) => acc + Number(p.In_Sum_Cost || 0), 0);

  return { daySales, dayPurch };
}

function prefillExpenseFormByDate(dateStr) {
  if (!dateStr) return;
  const { daySales, dayPurch } = computeDailyTotalsByDate(dateStr);
  const existing = expenseSummaryList.find(e => e.Ex_Date === dateStr);

  if (existing) {
    document.getElementById('expFormId').value = existing.Ex_ID;
    document.getElementById('expFormList').value = existing.Ex_List || `สรุปรายได้-ค่าใช้จ่ายประจำวัน (${dateStr})`;
    document.getElementById('expFormIncome').value = daySales > 0 ? daySales : Number(existing.Ex_Income || 0);
    document.getElementById('expFormPurch').value = dayPurch > 0 ? dayPurch : Number(existing.Ex_Purch || 0);
    document.getElementById('expFormOther').value = Number(existing.Ex_Other || 0);
  } else {
    document.getElementById('expFormId').value = '';
    document.getElementById('expFormList').value = `สรุปรายได้-ค่าใช้จ่ายประจำวัน (${dateStr})`;
    document.getElementById('expFormIncome').value = daySales;
    document.getElementById('expFormPurch').value = dayPurch;
    document.getElementById('expFormOther').value = 0;
  }
  previewExpenseFormProfit();
}

function previewExpenseFormProfit() {
  const inc = Number(document.getElementById('expFormIncome')?.value || 0);
  const pur = Number(document.getElementById('expFormPurch')?.value || 0);
  const oth = Number(document.getElementById('expFormOther')?.value || 0);
  const proLoss = inc - pur - oth;
  const previewEl = document.getElementById('expFormProfitPreview');
  if (previewEl) {
    previewEl.value = `${proLoss >= 0 ? '+' : ''}${formatMoney(proLoss)} บาท`;
    previewEl.style.color = proLoss >= 0 ? 'var(--color-success)' : 'var(--color-danger)';
  }
}

function resetExpenseForm() {
  const today = getTodayStr();
  document.getElementById('expFormId').value = '';
  document.getElementById('expFormDate').value = today;
  prefillExpenseFormByDate(today);
}

async function handleSaveExpenseForm(e) {
  e.preventDefault();
  const exId = document.getElementById('expFormId').value;
  const exDate = document.getElementById('expFormDate').value || getTodayStr();
  const exList = document.getElementById('expFormList').value.trim();
  const exIncome = Number(document.getElementById('expFormIncome').value || 0);
  const exPurch = Number(document.getElementById('expFormPurch').value || 0);
  const exOther = Number(document.getElementById('expFormOther').value || 0);
  const exProLoss = exIncome - exPurch - exOther;

  const payload = {
    Ex_Date: exDate,
    Ex_List: exList,
    Ex_Income: exIncome,
    Ex_Purch: exPurch,
    Ex_Other: exOther,
    Ex_Pro_Loss: exProLoss
  };

  if (exId) {
    const { error } = await db.from('Summary_Expans_TB').update(payload).eq('Ex_ID', Number(exId));
    if (error) {
      showToast(`เกิดข้อผิดพลาดในการอัปเดต: ${error.message}`, 'error');
      return;
    }
    showToast(`อัปเดตข้อมูลรายได้-ค่าใช้จ่ายวันที่ ${exDate} ลงตาราง Summary_Expans_TB สำเร็จ!`, 'success');
  } else {
    const { error } = await db.from('Summary_Expans_TB').insert([payload]);
    if (error) {
      showToast(`เกิดข้อผิดพลาดในการบันทึก: ${error.message}`, 'error');
      return;
    }
    showToast(`บันทึกข้อมูลค่าใช้จ่ายอื่นๆ (฿${formatMoney(exOther)}) ลงตาราง Summary_Expans_TB สำเร็จ!`, 'success');
  }

  await loadExpensesFromSupabase();
  resetExpenseForm();
}

async function syncDailyIncomeExpenseToSupabase(notify = true) {
  // Collect all unique dates from sales and purchases
  const dateSet = new Set();
  salesSummaryList.forEach(s => { if (s.S_Date) dateSet.add(s.S_Date); });
  purchaseSummaryList.forEach(p => { if (p.In_Date) dateSet.add(p.In_Date); });
  if (dateSet.size === 0) dateSet.add(getTodayStr());

  for (const d of dateSet) {
    const { daySales, dayPurch } = computeDailyTotalsByDate(d);
    const existing = expenseSummaryList.find(e => e.Ex_Date === d);

    if (existing) {
      const otherVal = Number(existing.Ex_Other || 0);
      const proLoss = daySales - dayPurch - otherVal;
      await db.from('Summary_Expans_TB').update({
        Ex_Income: daySales,
        Ex_Purch: dayPurch,
        Ex_Pro_Loss: proLoss
      }).eq('Ex_ID', existing.Ex_ID);
    } else {
      await db.from('Summary_Expans_TB').insert([{
        Ex_Date: d,
        Ex_List: `สรุปยอดรายได้-ค่าใช้จ่ายประจำวัน (${d})`,
        Ex_Income: daySales,
        Ex_Purch: dayPurch,
        Ex_Other: 0,
        Ex_Pro_Loss: daySales - dayPurch
      }]);
    }
  }

  await loadExpensesFromSupabase();
  if (notify) {
    showToast('วิเคราะห์ข้อมูลจากตารางคลังสินค้า/ซื้อ/ขาย และซิงค์ลงตาราง Summary_Expans_TB เรียบร้อยแล้ว', 'success');
  }
}

async function updateInlineOtherExpense(exId, newOtherVal) {
  const row = expenseSummaryList.find(e => e.Ex_ID === exId);
  if (!row) return;
  const other = Number(newOtherVal || 0);
  const inc = Number(row.Ex_Income || 0);
  const pur = Number(row.Ex_Purch || 0);
  const proLoss = inc - pur - other;

  const { error } = await db.from('Summary_Expans_TB').update({
    Ex_Other: other,
    Ex_Pro_Loss: proLoss
  }).eq('Ex_ID', exId);

  if (!error) {
    await loadExpensesFromSupabase();
    showToast(`อัปเดตค่าใช้จ่ายอื่นๆ (Ex_Other = ฿${formatMoney(other)}) เรียบร้อยแล้ว`, 'success');
  }
}

function editExpenseRowInForm(exId) {
  const row = expenseSummaryList.find(e => e.Ex_ID === exId);
  if (!row) return;
  document.getElementById('expFormId').value = row.Ex_ID;
  document.getElementById('expFormDate').value = row.Ex_Date || getTodayStr();
  document.getElementById('expFormList').value = row.Ex_List || '';
  document.getElementById('expFormIncome').value = Number(row.Ex_Income || 0);
  document.getElementById('expFormPurch').value = Number(row.Ex_Purch || 0);
  document.getElementById('expFormOther').value = Number(row.Ex_Other || 0);
  previewExpenseFormProfit();
  window.scrollTo({ top: 320, behavior: 'smooth' });
}

async function deleteExpenseRow(exId) {
  await db.from('Summary_Expans_TB').delete().eq('Ex_ID', exId);
  await loadExpensesFromSupabase();
  showToast('ลบรายการใน Summary_Expans_TB เรียบร้อยแล้ว', 'success');
}

function renderExpenseTable() {
  const tbody = document.getElementById('expenseTableBody');
  if (!tbody) return;

  let totalInc = 0;
  let totalPur = 0;
  let totalOth = 0;
  let totalNet = 0;

  if (expenseSummaryList.length === 0) {
    tbody.innerHTML = `<tr><td colspan="8" style="text-align:center; color:var(--color-muted);">ยังไม่มีข้อมูลใน Summary_Expans_TB (คลิกปุ่มวิเคราะห์และซิงค์ข้อมูล หรือกรอกแบบฟอร์มด้านบน)</td></tr>`;
  } else {
    tbody.innerHTML = expenseSummaryList.map(row => {
      const inc = Number(row.Ex_Income || 0);
      const pur = Number(row.Ex_Purch || 0);
      const oth = Number(row.Ex_Other || 0);
      const net = inc - pur - oth;

      totalInc += inc;
      totalPur += pur;
      totalOth += oth;
      totalNet += net;

      const isProfit = net >= 0;

      return `
        <tr>
          <td><strong>${row.Ex_ID}</strong></td>
          <td>${row.Ex_List || '-'}</td>
          <td style="color:var(--color-success); font-weight:600;">฿${formatMoney(inc)}</td>
          <td style="color:var(--color-secondary); font-weight:600;">฿${formatMoney(pur)}</td>
          <td>
            <input type="number" class="form-control" style="width:130px; padding:5px 8px; border-color:var(--color-accent); font-weight:700;"
              value="${oth}" step="0.01" onchange="updateInlineOtherExpense(${row.Ex_ID}, this.value)" />
          </td>
          <td style="font-weight:700; font-size:15px; color:${isProfit ? 'var(--color-success)' : 'var(--color-danger)'};">
            ${isProfit ? '+' : ''}฿${formatMoney(net)}
          </td>
          <td><strong>${row.Ex_Date || '-'}</strong></td>
          <td>
            <div style="display:flex; gap:6px;">
              <button class="btn btn-sm btn-outline" onclick="editExpenseRowInForm(${row.Ex_ID})">✏️ แก้ไข</button>
              <button class="btn btn-sm btn-danger" onclick="deleteExpenseRow(${row.Ex_ID})">🗑️ ลบ</button>
            </div>
          </td>
        </tr>
      `;
    }).join('');
  }

  document.getElementById('expStatIncome').textContent = `฿${formatMoney(totalInc)}`;
  document.getElementById('expStatPurch').textContent = `฿${formatMoney(totalPur)}`;
  document.getElementById('expStatOther').textContent = `฿${formatMoney(totalOth)}`;
  document.getElementById('expStatNetProfit').textContent = `${totalNet >= 0 ? '+' : ''}฿${formatMoney(totalNet)}`;
}

// ============================================================================
// PAGE 7: MONTHLY PROFIT & LOSS STATEMENT (งบกำไรขาดทุนรายเดือน) — ADMIN ONLY
// ============================================================================
function renderProfitLossPage() {
  const tbody = document.getElementById('plMonthlyTableBody');
  if (!tbody) return;

  const monthFilter = document.getElementById('plMonthFilter')?.value || '';
  const monthlyMap = {};

  const ensureMonth = ym => {
    if (!monthlyMap[ym]) {
      monthlyMap[ym] = { sales: 0, purchases: 0, otherExp: 0 };
    }
  };

  salesSummaryList.forEach(s => {
    const ym = (s.S_Date || getTodayStr()).slice(0, 7);
    ensureMonth(ym);
    monthlyMap[ym].sales += Number(s.S_Sum || 0);
  });

  purchaseSummaryList.forEach(p => {
    const ym = (p.In_Date || getTodayStr()).slice(0, 7);
    ensureMonth(ym);
    monthlyMap[ym].purchases += Number(p.In_Sum_Cost || 0);
  });

  expenseSummaryList.forEach(e => {
    const ym = (e.Ex_Date || getTodayStr()).slice(0, 7);
    ensureMonth(ym);
    monthlyMap[ym].otherExp += Number(e.Ex_Other || 0);
  });

  if (Object.keys(monthlyMap).length === 0) {
    ensureMonth(getTodayStr().slice(0, 7));
  }

  const months = Object.keys(monthlyMap)
    .filter(ym => !monthFilter || ym === monthFilter)
    .sort()
    .reverse();

  let grandSales = 0;
  let grandPurchases = 0;
  let grandOther = 0;

  tbody.innerHTML = months.map(ym => {
    const m = monthlyMap[ym];
    const grossProfit = m.sales - m.purchases;
    const netProfit = grossProfit - m.otherExp;
    grandSales += m.sales;
    grandPurchases += m.purchases;
    grandOther += m.otherExp;

    return `
      <tr>
        <td><strong>📅 ${ym}</strong></td>
        <td style="color:var(--color-success); font-weight:600;">฿${formatMoney(m.sales)}</td>
        <td style="color:var(--color-secondary); font-weight:600;">฿${formatMoney(m.purchases)}</td>
        <td style="font-weight:600;">฿${formatMoney(grossProfit)}</td>
        <td style="color:var(--color-accent); font-weight:600;">฿${formatMoney(m.otherExp)}</td>
        <td style="font-weight:700; font-size:15px; color:${netProfit >= 0 ? 'var(--color-success)' : 'var(--color-danger)'};">
          ${netProfit >= 0 ? '+' : ''}฿${formatMoney(netProfit)}
        </td>
        <td>
          <span class="badge ${netProfit >= 0 ? 'badge-ready' : 'badge-order'}">
            ${netProfit >= 0 ? 'กำไรสุทธิ' : 'ขาดทุนสุทธิ'}
          </span>
        </td>
      </tr>
    `;
  }).join('');

  const grandGross = grandSales - grandPurchases;
  const grandNet = grandGross - grandOther;

  document.getElementById('plTotalSales').textContent = `฿${formatMoney(grandSales)}`;
  document.getElementById('plTotalPurchases').textContent = `฿${formatMoney(grandPurchases)}`;
  document.getElementById('plTotalOtherExp').textContent = `฿${formatMoney(grandOther)}`;
  document.getElementById('plNetProfit').textContent = `${grandNet >= 0 ? '+' : ''}฿${formatMoney(grandNet)}`;

  document.getElementById('formalPlPeriodText').textContent = monthFilter
    ? `สำหรับงวดประจำเดือน ${monthFilter}`
    : 'สำหรับงวดสะสมทุกเดือน';

  document.getElementById('formalPlBody').innerHTML = `
    <div style="display:flex; justify-content:space-between;">
      <span><strong>รายได้จากการขายสินค้า (Sales Revenue)</strong></span>
      <strong>฿${formatMoney(grandSales)}</strong>
    </div>
    <div style="display:flex; justify-content:space-between; color:var(--color-secondary);">
      <span>หัก: ต้นทุนการซื้อสินค้า (Cost of Goods Purchased)</span>
      <span>(฿${formatMoney(grandPurchases)})</span>
    </div>
    <div style="display:flex; justify-content:space-between; border-top:1px solid var(--color-border); padding-top:4px;">
      <span><strong>กำไรขั้นต้น (Gross Profit)</strong></span>
      <strong>฿${formatMoney(grandGross)}</strong>
    </div>
    <div style="display:flex; justify-content:space-between; color:var(--color-accent);">
      <span>หัก: ค่าใช้จ่ายในการดำเนินงานอื่นๆ (Other Operating Expenses)</span>
      <span>(฿${formatMoney(grandOther)})</span>
    </div>
    <div style="display:flex; justify-content:space-between; border-top:2px solid var(--color-primary); border-bottom:3px double var(--color-primary); padding:8px 0; margin-top:8px; font-size:18px; color:${grandNet >= 0 ? 'var(--color-success)' : 'var(--color-danger)'};">
      <strong>กำไร (ขาดทุน) สุทธิประจำงวด (Net Profit / Loss)</strong>
      <strong>${grandNet >= 0 ? '+' : ''}฿${formatMoney(grandNet)}</strong>
    </div>
  `;
}

// ============================================================================
// PAGE 8: MEMBER MANAGEMENT & REGISTRATION BY ADMIN (Employee_TB)
// Strict validations:
// 1. ชื่อ-นามสกุล: ห้ามใส่ชื่อหรือนามสกุลมั่ว หรือไม่มีตามบัญญัติในพจนานุกรม
// 2. เบอร์โทร: ฟิกไว้ที่ 10 ตัวอักษรตัวเลขเท่านั้น ห้ามมีขีดหรืออักขระพิเศษใดๆ ทั้งสิ้น
// 3. อีเมล: ตรวจสอบองค์ประกอบอีเมลให้ครบ
// 4. รหัสผู้ใช้งาน: ประมวลผลให้อัตโนมัติ (AP-1122x)
// ============================================================================
function validateDictionaryThaiOrEnglishName(fullName) {
  const clean = (fullName || '').trim().replace(/\s+/g, ' ');
  const parts = clean.split(' ');

  if (parts.length < 2) {
    return { valid: false, reason: 'กรุณาระบุทั้งชื่อจริงและนามสกุล คั่นด้วยช่องว่าง (เช่น กัญญาณี ศรีสวัสดิ์)' };
  }

  // Keyboard smash / gibberish blacklist
  const gibberishPatterns = [
    /asdf/i, /qwer/i, /zxcv/i, /abcd/i, /test/i, /xxxx/i, /aaaa/i,
    /ฟหกด/, /่าสว/, /ผปแอ/, /กขคง/, /ดดด/, /กกก/, /ขขข/, /อออ/, /มมม/,
    /(.)\1{2,}/ // 3+ identical consecutive chars
  ];

  for (const pat of gibberishPatterns) {
    if (pat.test(clean)) {
      return { valid: false, reason: 'ตรวจพบการพิมพ์ชื่อหรือนามสกุลมั่ว หรือไม่อยู่ตามหลักพจนานุกรม' };
    }
  }

  for (const part of parts) {
    if (part.length < 2 || part.length > 35) {
      return { valid: false, reason: 'ความยาวของชื่อและนามสกุลต้องมีความยาวอย่างน้อย 2 ตัวอักษรขึ้นไป' };
    }

    const isThai = /^[\u0E01-\u0E4E]+$/.test(part);
    const isEnglish = /^[A-Za-z]+$/.test(part);

    if (!isThai && !isEnglish) {
      return { valid: false, reason: 'ชื่อและนามสกุลต้องเป็นตัวอักษรภาษาไทยหรือภาษาอังกฤษเท่านั้น ห้ามมีตัวเลขหรือสัญลักษณ์' };
    }

    if (isThai) {
      // Must contain at least one Thai consonant (\u0E01-\u0E2E) AND at least one Thai vowel/tone or valid word structure
      const hasConsonant = /[\u0E01-\u0E2E]/.test(part);
      const hasVowelOrMedial = /[\u0E30-\u0E3A\u0E40-\u0E47\u0E2F\u0E4C\u0E2D\u0E27\u0E22\u0E23\u0E19\u0E21\u0E07]/.test(part);
      // Cannot start with combining vowel/tone mark
      const startsWithCombining = /^[\u0E30-\u0E3A\u0E47-\u0E4E]/.test(part);
      if (!hasConsonant || !hasVowelOrMedial || startsWithCombining) {
        return { valid: false, reason: `คำว่า "${part}" ไม่ถูกต้องตามโครงสร้างคำในพจนานุกรมภาษาไทย (ต้องประกอบด้วยพยัญชนะและสระที่ถูกต้อง)` };
      }
    }

    if (isEnglish) {
      const hasEngVowel = /[aeiouyAEIOUY]/.test(part);
      const hasEngConsonant = /[bcdfghjklmnpqrstvwxzBCDFGHJKLMNPQRSTVWXZ]/.test(part);
      if (!hasEngVowel || !hasEngConsonant) {
        return { valid: false, reason: `"${part}" ไม่ถูกต้องตามหลักคำศัพท์พจนานุกรม (ต้องมีทั้งพยัญชนะและสระ)` };
      }
    }
  }

  return { valid: true };
}

function validateStrictPhone10Digits(phone) {
  if (!/^[0-9]{10}$/.test(phone)) {
    return { valid: false, reason: 'เบอร์โทรศัพท์ต้องเป็นตัวเลข 10 หลักเท่านั้น (เช่น 0812345678) ห้ามมีขีด (-) หรืออักขระพิเศษใดๆ ทั้งสิ้น' };
  }
  if (!/^0[2689][0-9]{8}$/.test(phone)) {
    return { valid: false, reason: 'เบอร์โทรศัพท์ต้องขึ้นต้นด้วยรหัสหมายเลขโทรศัพท์ที่ถูกต้อง (เช่น 06, 08, 09, 02) ครบ 10 หลัก' };
  }
  return { valid: true };
}

function validateCompleteEmail(email) {
  const emailRegex = /^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$/;
  if (!emailRegex.test(email)) {
    return { valid: false, reason: 'รูปแบบอีเมลไม่ครบองค์ประกอบ (ต้องประกอบด้วยชื่อผู้ใช้, เครื่องหมาย @, ชื่อโดเมน และนามสกุลโดเมน เช่น .com, .co.th)' };
  }
  return { valid: true };
}

function generateNextMemberCode() {
  const mode = document.getElementById('memberEditMode')?.value;
  if (mode === 'edit') return;

  let maxNum = 11220;
  memberList.forEach(m => {
    const match = (m.Em_ID || '').match(/(\d+)$/);
    if (match) {
      const n = parseInt(match[1], 10);
      if (n > maxNum) maxNum = n;
    }
  });
  const nextCode = `AP-${maxNum + 1}`;
  const input = document.getElementById('memIdInput');
  if (input) input.value = nextCode;
}

function resetMemberForm() {
  document.getElementById('memberEditMode').value = 'create';
  document.getElementById('memberFormTitle').textContent = '➕ เพิ่มผู้ใช้งานใหม่ / ลงทะเบียนสมาชิกโดยแอดมิน';
  document.getElementById('memNameInput').value = '';
  document.getElementById('memTelInput').value = '';
  document.getElementById('memEmailInput').value = '';
  document.getElementById('memRoleDescInput').value = 'พนักงานขายหน้าร้าน แผนกเครื่องสำอาง';
  document.getElementById('memStatusInput').value = 'active';
  generateNextMemberCode();
}

function editMemberRow(emId) {
  const m = memberList.find(x => x.Em_ID === emId);
  if (!m) return;
  document.getElementById('memberEditMode').value = 'edit';
  document.getElementById('memberFormTitle').textContent = `✏️ แก้ไขข้อมูลสมาชิก (${m.Em_ID})`;
  document.getElementById('memNameInput').value = m.Em_Name || '';
  document.getElementById('memTelInput').value = m.EM_Tell || '';
  document.getElementById('memEmailInput').value = m.EM_Email || '';
  document.getElementById('memIdInput').value = m.Em_ID;
  document.getElementById('memRoleDescInput').value = m.Em_User || '';
  document.getElementById('memStatusInput').value = m.Em_Status || 'active';
  window.scrollTo({ top: 300, behavior: 'smooth' });
}

async function handleSaveMember(e) {
  e.preventDefault();
  const mode = document.getElementById('memberEditMode').value;
  const fullName = document.getElementById('memNameInput').value.trim().replace(/\s+/g, ' ');
  const phone = document.getElementById('memTelInput').value.trim();
  const email = document.getElementById('memEmailInput').value.trim().toLowerCase();
  const emId = document.getElementById('memIdInput').value.trim();
  const roleDesc = document.getElementById('memRoleDescInput').value.trim() || 'พนักงานร้านเครื่องสำอาง AP';
  const status = document.getElementById('memStatusInput').value;

  // 1. Validate Name
  const nameCheck = validateDictionaryThaiOrEnglishName(fullName);
  if (!nameCheck.valid) {
    showToast(nameCheck.reason, 'error');
    return;
  }

  // 2. Validate Phone (10 digits only, no dashes/special chars)
  const phoneCheck = validateStrictPhone10Digits(phone);
  if (!phoneCheck.valid) {
    showToast(phoneCheck.reason, 'error');
    return;
  }

  // 3. Validate Email
  const emailCheck = validateCompleteEmail(email);
  if (!emailCheck.valid) {
    showToast(emailCheck.reason, 'error');
    return;
  }

  const payload = {
    Em_ID: emId,
    Em_Name: fullName,
    EM_Tell: phone,
    EM_Email: email,
    EM_date: getTodayStr(),
    Em_Status: status,
    Em_User: roleDesc
  };

  if (mode === 'create') {
    const { error } = await db.from('Employee_TB').insert([payload]);
    if (error) {
      showToast(`เกิดข้อผิดพลาดในการเพิ่มสมาชิก: ${error.message}`, 'error');
      return;
    }
    showToast(`เพิ่มผู้ใช้งาน ${fullName} (รหัสอัตโนมัติ: ${emId}) ลงใน Employee_TB สำเร็จ!`, 'success');
  } else {
    const { error } = await db.from('Employee_TB').update(payload).eq('Em_ID', emId);
    if (error) {
      showToast(`เกิดข้อผิดพลาดในการแก้ไขสมาชิก: ${error.message}`, 'error');
      return;
    }
    showToast(`อัปเดตข้อมูลสมาชิก ${fullName} (${emId}) สำเร็จ!`, 'success');
  }

  await loadMembersFromSupabase();
  resetMemberForm();
}

async function deleteMemberRow(emId, name) {
  const { error } = await db.from('Employee_TB').delete().eq('Em_ID', emId);
  if (error) {
    showToast(`ลบสมาชิกไม่สำเร็จ: ${error.message}`, 'error');
    return;
  }
  await loadMembersFromSupabase();
  showToast(`ลบรายชื่อสมาชิก ${name} (${emId}) ออกจากระบบเรียบร้อยแล้ว`, 'success');
}

function renderMemberTable() {
  const tbody = document.getElementById('memberTableBody');
  const badge = document.getElementById('memberTotalBadge');
  if (!tbody) return;
  if (badge) badge.textContent = `${memberList.length} คน`;

  if (memberList.length === 0) {
    tbody.innerHTML = `<tr><td colspan="8" style="text-align:center; color:var(--color-muted);">ยังไม่มีรายชื่อสมาชิกในระบบ</td></tr>`;
    return;
  }

  tbody.innerHTML = memberList.map(m => `
    <tr>
      <td><span class="shelf-bar-tag">${m.Em_ID}</span></td>
      <td><strong>${m.Em_Name}</strong></td>
      <td>${m.EM_Tell || '-'}</td>
      <td>${m.EM_Email}</td>
      <td>${m.EM_date || '-'}</td>
      <td>
        <span class="badge ${m.Em_Status === 'active' ? 'badge-ready' : 'badge-order'}">
          ${m.Em_Status || 'active'}
        </span>
      </td>
      <td>${m.Em_User || '-'}</td>
      <td>
        <div style="display:flex; gap:6px;">
          <button class="btn btn-sm btn-outline" onclick="editMemberRow('${m.Em_ID}')">✏️ แก้ไข</button>
          <button class="btn btn-sm btn-danger" onclick="deleteMemberRow('${m.Em_ID}', '${m.Em_Name}')">🗑️ ลบ</button>
        </div>
      </td>
    </tr>
  `).join('');
}

// ============================================================================
// REALTIME MULTI-DEVICE SUBSCRIPTION
// ("ให้ข้อมูลทุกข้อมูลอัปเดตให้เป็นปัจจุบันและให้ทุกเครื่องสามารถมองเห็นข้อมูลถึงกันได้ โดยจะใช้ฐานข้อมูลเดียวกันคือ APSHOP1122'sOrg")
// ============================================================================
function setupRealtimeSubscriptions() {
  const channel = db
    .channel('apshop1122-realtime-sync')
    .on('postgres_changes', { event: '*', schema: 'public', table: 'Inventory_TB' }, () => {
      loadInventoryFromSupabase();
      loadPurchaseSelectFromSupabase();
    })
    .on('postgres_changes', { event: '*', schema: 'public', table: 'Inventorys_TB' }, () => {
      loadInventoryFromSupabase();
    })
    .on('postgres_changes', { event: '*', schema: 'public', table: 'Summary_Purch_TB' }, () => {
      loadPurchaseSummaryFromSupabase();
      loadInventoryFromSupabase();
    })
    .on('postgres_changes', { event: '*', schema: 'public', table: 'Summary_Self_TB' }, () => {
      loadSalesSummaryFromSupabase();
      loadInventoryFromSupabase();
    })
    .on('postgres_changes', { event: '*', schema: 'public', table: 'Summary_Expans_TB' }, () => {
      loadExpensesFromSupabase();
    })
    .on('postgres_changes', { event: '*', schema: 'public', table: 'Employee_TB' }, () => {
      loadMembersFromSupabase();
    })
    .subscribe(status => {
      const dot = document.getElementById('rtStatusDot');
      const txt = document.getElementById('rtStatusText');
      if (status === 'SUBSCRIBED') {
        if (dot) dot.style.background = '#22C55E';
        if (txt) txt.textContent = "ออนไลน์: ซิงค์ฐานข้อมูล APSHOP1122'sOrg แบบ Realtime ทุกเครื่อง";
      }
    });
  return channel;
}

// ============================================================================
// INITIALIZATION ON DOM LOAD
// ============================================================================
window.addEventListener('DOMContentLoaded', async () => {
  const today = getTodayStr();
  ['purchDateInput', 'salesDateInput', 'expFormDate'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = today;
  });

  await syncAllData(false);
  resetExpenseForm();
  setupRealtimeSubscriptions();
});
