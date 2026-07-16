import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

const sb = createClient(window.SUPABASE_URL, window.SUPABASE_ANON_KEY);

/* ============================================================
   Utilities
   ============================================================ */
const $ = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));
const uid = () => crypto.randomUUID();

function fmt(n) {
  const v = Number(n) || 0;
  return v.toLocaleString("th-TH", { minimumFractionDigits: 0, maximumFractionDigits: 2 });
}

function toast(msg, isError = false) {
  const el = $("#toast");
  el.textContent = msg;
  el.classList.toggle("error", isError);
  el.classList.add("show");
  clearTimeout(toast._t);
  toast._t = setTimeout(() => el.classList.remove("show"), 2600);
}

// แปลงตัวเลขเป็นคำอ่านภาษาไทย (บาทถ้วน) — ใช้แสดงยอดรวมแบบเดียวกับฟอร์มต้นฉบับ
function thaiBahtText(number) {
  const numText = ["ศูนย์","หนึ่ง","สอง","สาม","สี่","ห้า","หก","เจ็ด","แปด","เก้า"];
  const unitText = ["","สิบ","ร้อย","พัน","หมื่น","แสน","ล้าน"];
  number = Math.abs(Number(number) || 0);
  const [intPartStr, decPartStr = "0"] = number.toFixed(2).split(".");

  function readNumber(numStr) {
    numStr = numStr.replace(/^0+/, "");
    if (numStr === "") return "";
    let result = "";
    const len = numStr.length;
    for (let i = 0; i < len; i++) {
      const digit = parseInt(numStr[i], 10);
      const place = len - i - 1;
      if (digit === 0) continue;
      if (place % 6 === 1 && digit === 1) {
        result += "สิบ";
      } else if (place % 6 === 1 && digit === 2) {
        result += "ยี่สิบ";
      } else if (place % 6 === 0 && digit === 1 && len > 1 && place !== len - 1) {
        result += "เอ็ด";
      } else {
        result += numText[digit] + unitText[place % 6];
      }
      if (place > 0 && place % 6 === 0 && place < len - 1) result += "ล้าน";
    }
    return result;
  }

  const intPart = readNumber(intPartStr) || "ศูนย์";
  const decPart = readNumber(decPartStr);
  if (decPartStr === "00") return `${intPart}บาทถ้วน`;
  return `${intPart}บาท${decPart}สตางค์`;
}

function todayISO() {
  return new Date().toISOString().slice(0, 10);
}

/* ============================================================
   State
   ============================================================ */
const state = {
  tab: "adv",           // 'adv' = FM-AC-05, 'exp' = FM-AC-06
  session: null,
  advForm: null,        // header object being edited (null id = new)
  advItems: [],
  advRecords: [],
  advSearch: "",
  expForm: null,
  expItems: [],
  expRecords: [],
  expSearch: "",
};

const ADV_COST_FIELDS = [
  { key: "cost_phone",      label: "ค่าโทรศัพท์" },
  { key: "cost_travel",     label: "ค่าเดินทาง" },
  { key: "cost_postage",    label: "ค่าไปรษณีย์" },
  { key: "cost_taxi",       label: "ค่าแท็กซี่" },
  { key: "cost_fuel",       label: "ค่าน้ำมัน" },
  { key: "cost_tollway",    label: "ค่าทางด่วน" },
  { key: "cost_food",       label: "ค่าอาหาร" },
  { key: "cost_hotel",      label: "ค่าที่พัก" },
  { key: "cost_stationery", label: "เครื่องเขียนแบบพิมพ์" },
  { key: "cost_other",      label: "อื่นๆ" },
];

function blankAdvForm() {
  return {
    id: null, employee_name: "", department: "", project: "", province: "",
    doc_year: new Date().getFullYear() + 543, adv_amount: 0, adv_ref: "",
    transfer_to_company_date: "", company_return_to_employee_date: "",
    transfer_proof_note: "", requester_name: "", approver_name: "",
    receiver_name: "", checker_name: "", accountant_name: "", status: "draft",
  };
}
function blankAdvItem() {
  const item = { _uid: uid(), item_month: "", item_date: todayISO(), description: "", clear_amount: 0, has_bill: true };
  ADV_COST_FIELDS.forEach(f => item[f.key] = 0);
  item._type = "cost_travel";
  return item;
}
function blankExpForm() {
  return { id: null, employee_name: "", position: "", department: "", project: "", requester_name: "", approver_name: "", status: "draft" };
}
function blankExpItem() {
  return { _uid: uid(), seq: 1, item_date: todayISO(), description: "", amount: 0, travel_start: "", travel_end: "", guest_count: "", guest_names: "", other_note: "", remark: "" };
}

/* ============================================================
   Auth
   ============================================================ */
async function initAuth() {
  const { data } = await sb.auth.getSession();
  state.session = data.session;
  sb.auth.onAuthStateChange((_event, session) => {
    state.session = session;
    renderRoot();
  });
  renderRoot();
}

async function handleLoginSubmit(e) {
  e.preventDefault();
  const email = $("#login-email").value.trim();
  const password = $("#login-password").value;
  const errEl = $("#login-error");
  errEl.textContent = "";
  const { error } = await sb.auth.signInWithPassword({ email, password });
  if (error) {
    errEl.textContent = "เข้าสู่ระบบไม่สำเร็จ: " + error.message;
    return;
  }
}

async function handleLogout() {
  await sb.auth.signOut();
}

/* ============================================================
   Root render
   ============================================================ */
function renderRoot() {
  const root = $("#app-root");
  if (!state.session) {
    root.innerHTML = renderLogin();
    $("#login-form").addEventListener("submit", handleLoginSubmit);
    return;
  }
  root.innerHTML = renderShell();
  $("#logout-btn").addEventListener("click", handleLogout);
  $$(".tab-btn").forEach(btn => btn.addEventListener("click", () => {
    state.tab = btn.dataset.tab;
    renderRoot();
  }));
  if (state.tab === "adv") {
    if (!state.advForm) newAdvForm(false);
    mountAdvTab();
  } else {
    if (!state.expForm) newExpForm(false);
    mountExpTab();
  }
}

function renderLogin() {
  return `
  <div class="login-wrap">
    <div class="login-card">
      <span class="eyebrow">FM-AC-05 / FM-AC-06</span>
      <h1>ระบบเคลียร์เงินทดรองจ่าย</h1>
      <form id="login-form">
        <div class="field">
          <label>อีเมล</label>
          <input id="login-email" type="email" required autocomplete="username" />
        </div>
        <div class="field">
          <label>รหัสผ่าน</label>
          <input id="login-password" type="password" required autocomplete="current-password" />
        </div>
        <button class="btn btn-primary" type="submit">เข้าสู่ระบบ</button>
        <div class="login-error" id="login-error"></div>
      </form>
    </div>
  </div>`;
}

function renderShell() {
  return `
  <div class="topbar no-print">
    <div class="brand">
      <small>ระบบเบิก–เคลียร์เงินทดรองจ่าย</small>
      <h1>FM-AC-05 / FM-AC-06</h1>
    </div>
    <div class="user">
      <span>${state.session.user.email}</span>
      <button id="logout-btn">ออกจากระบบ</button>
    </div>
  </div>
  <div class="tabbar no-print">
    <button class="tab-btn ${state.tab === "adv" ? "active" : ""}" data-tab="adv">FM-AC-05 เคลียร์เงินทดรองจ่าย</button>
    <button class="tab-btn ${state.tab === "exp" ? "active" : ""}" data-tab="exp">FM-AC-06 ค่าใช้จ่าย (ไม่มีเอกสาร)</button>
  </div>
  <div class="wrap" id="tab-content"></div>
  <div class="toast" id="toast"></div>`;
}

/* ============================================================
   FM-AC-05: ใบสรุปเคลียร์เงินทดรองจ่าย
   ============================================================ */
function newAdvForm(doRender = true) {
  state.advForm = blankAdvForm();
  state.advItems = [blankAdvItem()];
  if (doRender) mountAdvTab();
}

function advTotals() {
  const totalClear = state.advItems.reduce((s, it) => s + (Number(it.clear_amount) || 0), 0);
  const advAmount = Number(state.advForm.adv_amount) || 0;
  const remaining = advAmount - totalClear;
  const byType = {};
  ADV_COST_FIELDS.forEach(f => byType[f.key] = 0);
  state.advItems.forEach(it => {
    const t = it._type || "cost_other";
    byType[t] = (byType[t] || 0) + (Number(it.clear_amount) || 0);
  });
  return { totalClear, advAmount, remaining, byType };
}

function mountAdvTab() {
  const target = $("#tab-content");
  const f = state.advForm;
  const { totalClear, advAmount, remaining, byType } = advTotals();

  const balanceClass = remaining > 0 ? "over" : remaining < 0 ? "short" : "match";
  const balanceLabel = remaining > 0
    ? "พนักงานต้องคืนเงินบริษัท"
    : remaining < 0
      ? "บริษัทต้องคืนเงินพนักงาน"
      : "ยอดตรงกันพอดี";

  target.innerHTML = `
    <div class="grid-2">
      <div>
        <div class="card">
          <h2><span class="eyebrow">ข้อมูลผู้เบิก</span>ใบสรุปเคลียร์เงินทดรองจ่าย</h2>
          <div class="field-grid">
            <div class="field"><label>ชื่อ-นามสกุล</label><input id="a-employee_name" value="${esc(f.employee_name)}" placeholder="เช่น นายพชรพงศ์ ธันวงศ์" /></div>
            <div class="field"><label>แผนก</label><input id="a-department" value="${esc(f.department)}" /></div>
            <div class="field span-2"><label>โครงการ</label><input id="a-project" value="${esc(f.project)}" placeholder="เช่น OnSite ติดตั้งระบบงานครุภัณฑ์ รพ...." /></div>
            <div class="field"><label>จังหวัด</label><input id="a-province" value="${esc(f.province)}" /></div>
            <div class="field"><label>ปี พ.ศ.</label><input id="a-doc_year" type="number" value="${f.doc_year}" /></div>
            <div class="field"><label>จำนวนเงิน Advance (บาท)</label><input id="a-adv_amount" type="number" step="0.01" value="${advAmount}" /></div>
            <div class="field"><label>อ้างอิง ADV</label><input id="a-adv_ref" value="${esc(f.adv_ref)}" /></div>
            <div class="field"><label>โอนเงินคืนบริษัทวันที่</label><input id="a-transfer_to_company_date" type="date" value="${f.transfer_to_company_date || ""}" /></div>
            <div class="field"><label>บริษัทฯ คืนเงินพนักงานวันที่</label><input id="a-company_return_to_employee_date" type="date" value="${f.company_return_to_employee_date || ""}" /></div>
            <div class="field span-2"><label>หลักฐานการโอนเงิน (เช่น เลขที่ Slip)</label><input id="a-transfer_proof_note" value="${esc(f.transfer_proof_note)}" /></div>
          </div>
        </div>

        <div class="card">
          <h2><span class="eyebrow">รายการค่าใช้จ่าย</span>รายการเคลียร์ค่าใช้จ่าย</h2>
          <div class="table-scroll">
            <table class="items" id="adv-items-table">
              <thead>
                <tr>
                  <th>เดือน</th><th>วันที่</th><th style="min-width:220px">รายการ</th>
                  <th style="min-width:130px">ประเภทค่าใช้จ่าย</th><th>มีบิล/ไม่มีบิล</th>
                  <th style="min-width:110px">จำนวนเงิน</th><th></th>
                </tr>
              </thead>
              <tbody>
                ${state.advItems.map((it, idx) => advItemRow(it, idx)).join("")}
              </tbody>
            </table>
          </div>
          <button class="add-row-btn no-print" id="adv-add-row">+ เพิ่มรายการ</button>
        </div>

        <div class="action-row no-print">
          <div class="card" style="flex:1; min-width:260px">
            <h2><span class="eyebrow">ผู้เกี่ยวข้อง</span>ผู้ขอเบิก / ผู้อนุมัติ / อื่นๆ</h2>
            <div class="field-grid">
              <div class="field"><label>ผู้ขอเบิก</label><input id="a-requester_name" value="${esc(f.requester_name)}" /></div>
              <div class="field"><label>ผู้อนุมัติ</label><input id="a-approver_name" value="${esc(f.approver_name)}" /></div>
              <div class="field"><label>ผู้รับเงิน</label><input id="a-receiver_name" value="${esc(f.receiver_name)}" /></div>
              <div class="field"><label>ผู้ตรวจสอบ</label><input id="a-checker_name" value="${esc(f.checker_name)}" /></div>
              <div class="field span-2"><label>ผู้บันทึกบัญชี</label><input id="a-accountant_name" value="${esc(f.accountant_name)}" /></div>
              <div class="field"><label>สถานะ</label>
                <select id="a-status">
                  ${["draft","submitted","approved"].map(s => `<option value="${s}" ${f.status===s?"selected":""}>${statusLabel(s)}</option>`).join("")}
                </select>
              </div>
            </div>
          </div>
        </div>
      </div>

      <div>
        <div class="ledger">
          <div class="ledger-head">
            <span class="eyebrow">สรุปยอด</span>
            <h3>ใบสรุปเคลียร์เงินทดรอง</h3>
          </div>
          <div class="ledger-perf"></div>
          <div class="ledger-body">
            <div class="ledger-line"><span>จำนวนเงิน Advance</span><span class="amt">฿${fmt(advAmount)}</span></div>
            <div class="ledger-line"><span>จำนวนเงินค่าใช้จ่าย</span><span class="amt">฿${fmt(totalClear)}</span></div>
            <div class="ledger-line"><span>คงเหลือ เกิน (+) / ขาด (-)</span><span class="amt">${remaining >= 0 ? "" : "-"}฿${fmt(Math.abs(remaining))}</span></div>
            <div class="balance-pill ${balanceClass}">${balanceLabel}</div>

            <div class="breakdown-title">แยกตามประเภทค่าใช้จ่าย</div>
            <div class="breakdown-grid">
              ${ADV_COST_FIELDS.map(fdef => `<div><span>${fdef.label}</span><span class="amt">฿${fmt(byType[fdef.key])}</span></div>`).join("")}
            </div>
          </div>
        </div>

        <div class="card no-print" style="margin-top:18px">
          <h2>บันทึกข้อมูล</h2>
          <div class="action-row">
            <button class="btn btn-primary" id="adv-save-btn">บันทึกลง Supabase</button>
            <button class="btn btn-ghost" id="adv-new-btn">สร้างใบใหม่</button>
            <button class="btn btn-ghost" id="adv-print-btn">พิมพ์ / PDF</button>
            ${f.id ? `<button class="btn btn-danger" id="adv-delete-btn">ลบใบนี้</button>` : ""}
          </div>
        </div>

        <div class="card no-print" style="margin-top:18px">
          <h2>รายการที่บันทึกไว้</h2>
          <div class="list-toolbar">
            <input id="adv-search" placeholder="ค้นหาชื่อ / โครงการ / อ้างอิง ADV" value="${esc(state.advSearch)}" />
          </div>
          <div id="adv-records-list">${renderAdvRecordsTable()}</div>
        </div>
      </div>
    </div>
  `;

  bindAdvEvents();
  loadAdvRecords();
}

function advItemRow(it, idx) {
  return `
  <tr data-uid="${it._uid}">
    <td><input class="a-item-field" data-field="item_month" data-idx="${idx}" value="${esc(it.item_month)}" placeholder="เช่น พ.ค." style="width:60px" /></td>
    <td><input class="a-item-field" data-field="item_date" data-idx="${idx}" type="date" value="${it.item_date || ""}" /></td>
    <td><input class="a-item-field" data-field="description" data-idx="${idx}" value="${esc(it.description)}" placeholder="รายละเอียดค่าใช้จ่าย" /></td>
    <td>
      <select class="a-item-field" data-field="_type" data-idx="${idx}">
        ${ADV_COST_FIELDS.map(f2 => `<option value="${f2.key}" ${it._type===f2.key?"selected":""}>${f2.label}</option>`).join("")}
      </select>
    </td>
    <td>
      <select class="a-item-field" data-field="has_bill" data-idx="${idx}">
        <option value="1" ${it.has_bill ? "selected" : ""}>มีบิล</option>
        <option value="0" ${!it.has_bill ? "selected" : ""}>ไม่มีบิล</option>
      </select>
    </td>
    <td><input class="a-item-field" data-field="clear_amount" data-idx="${idx}" type="number" step="0.01" value="${it.clear_amount}" /></td>
    <td class="center"><button class="row-del a-item-del" data-idx="${idx}" title="ลบแถว">✕</button></td>
  </tr>`;
}

function bindAdvEvents() {
  $("#adv-add-row").addEventListener("click", () => {
    state.advItems.push(blankAdvItem());
    mountAdvTab();
  });
  $$(".a-item-field").forEach(el => {
    el.addEventListener("input", () => {
      const idx = Number(el.dataset.idx);
      const field = el.dataset.field;
      let val = el.value;
      if (field === "has_bill") val = val === "1";
      state.advItems[idx][field] = val;
      // ตัวเลขต้อง re-render ยอดรวม ไม่ต้อง re-render ทั้งตาราง เพื่อไม่ให้ focus หลุด ยกเว้นช่องตัวเลข/ประเภท
      if (field === "clear_amount" || field === "_type") {
        refreshAdvSummaryOnly();
      }
    });
  });
  $$(".a-item-del").forEach(btn => {
    btn.addEventListener("click", () => {
      const idx = Number(btn.dataset.idx);
      state.advItems.splice(idx, 1);
      if (state.advItems.length === 0) state.advItems.push(blankAdvItem());
      mountAdvTab();
    });
  });

  const fieldMap = ["employee_name","department","project","province","doc_year","adv_amount","adv_ref",
    "transfer_to_company_date","company_return_to_employee_date","transfer_proof_note",
    "requester_name","approver_name","receiver_name","checker_name","accountant_name","status"];
  const advNumericFields = new Set(["doc_year", "adv_amount"]);
  fieldMap.forEach(key => {
    const el = $(`#a-${key}`);
    if (!el) return;
    el.addEventListener("input", () => {
      state.advForm[key] = advNumericFields.has(key) ? (Number(el.value) || 0) : el.value;
      if (key === "adv_amount") refreshAdvSummaryOnly();
    });
  });

  $("#adv-save-btn").addEventListener("click", saveAdvForm);
  $("#adv-new-btn").addEventListener("click", () => newAdvForm());
  $("#adv-print-btn").addEventListener("click", () => window.print());
  const delBtn = $("#adv-delete-btn");
  if (delBtn) delBtn.addEventListener("click", deleteAdvForm);
  $("#adv-search").addEventListener("input", (e) => {
    state.advSearch = e.target.value;
    $("#adv-records-list").innerHTML = renderAdvRecordsTable();
    bindAdvRecordClicks();
  });
  bindAdvRecordClicks();
}

function refreshAdvSummaryOnly() {
  const { totalClear, advAmount, remaining, byType } = advTotals();
  const balanceClass = remaining > 0 ? "over" : remaining < 0 ? "short" : "match";
  const balanceLabel = remaining > 0 ? "พนักงานต้องคืนเงินบริษัท" : remaining < 0 ? "บริษัทต้องคืนเงินพนักงาน" : "ยอดตรงกันพอดี";
  const ledgerBody = document.querySelector(".ledger-body");
  if (!ledgerBody) return;
  ledgerBody.innerHTML = `
    <div class="ledger-line"><span>จำนวนเงิน Advance</span><span class="amt">฿${fmt(advAmount)}</span></div>
    <div class="ledger-line"><span>จำนวนเงินค่าใช้จ่าย</span><span class="amt">฿${fmt(totalClear)}</span></div>
    <div class="ledger-line"><span>คงเหลือ เกิน (+) / ขาด (-)</span><span class="amt">${remaining >= 0 ? "" : "-"}฿${fmt(Math.abs(remaining))}</span></div>
    <div class="balance-pill ${balanceClass}">${balanceLabel}</div>
    <div class="breakdown-title">แยกตามประเภทค่าใช้จ่าย</div>
    <div class="breakdown-grid">
      ${ADV_COST_FIELDS.map(fdef => `<div><span>${fdef.label}</span><span class="amt">฿${fmt(byType[fdef.key])}</span></div>`).join("")}
    </div>`;
}

function renderAdvRecordsTable() {
  const q = state.advSearch.trim().toLowerCase();
  const rows = state.advRecords.filter(r =>
    !q || (r.employee_name || "").toLowerCase().includes(q) ||
    (r.project || "").toLowerCase().includes(q) ||
    (r.adv_ref || "").toLowerCase().includes(q));
  if (rows.length === 0) return `<div class="empty-state">ยังไม่มีรายการที่บันทึกไว้ — กรอกฟอร์มด้านบนแล้วกดบันทึก</div>`;
  return `
  <table class="records">
    <thead><tr><th>ชื่อ</th><th>โครงการ</th><th>Advance</th><th>สถานะ</th></tr></thead>
    <tbody>
      ${rows.map(r => `
        <tr data-id="${r.id}">
          <td>${esc(r.employee_name)}</td>
          <td>${esc(r.project || "-")}</td>
          <td>฿${fmt(r.adv_amount)}</td>
          <td><span class="status-tag ${r.status}">${statusLabel(r.status)}</span></td>
        </tr>`).join("")}
    </tbody>
  </table>`;
}

function bindAdvRecordClicks() {
  $$("#adv-records-list tr[data-id]").forEach(tr => {
    tr.addEventListener("click", () => loadAdvForm(tr.dataset.id));
  });
}

async function loadAdvRecords() {
  const { data, error } = await sb.from("adv_clear_forms").select("*").order("created_at", { ascending: false });
  if (error) { toast("โหลดรายการไม่สำเร็จ: " + error.message, true); return; }
  state.advRecords = data || [];
  const listEl = $("#adv-records-list");
  if (listEl) { listEl.innerHTML = renderAdvRecordsTable(); bindAdvRecordClicks(); }
}

async function loadAdvForm(id) {
  const { data: form, error } = await sb.from("adv_clear_forms").select("*").eq("id", id).single();
  if (error) { toast("โหลดข้อมูลไม่สำเร็จ: " + error.message, true); return; }
  const { data: items, error: itemErr } = await sb.from("adv_clear_items").select("*").eq("form_id", id).order("sort_order");
  if (itemErr) { toast("โหลดรายการค่าใช้จ่ายไม่สำเร็จ: " + itemErr.message, true); return; }

  state.advForm = form;
  state.advItems = (items || []).map(it => {
    const type = ADV_COST_FIELDS.find(f => Number(it[f.key]) > 0)?.key || "cost_other";
    return { ...it, _uid: uid(), _type: type };
  });
  if (state.advItems.length === 0) state.advItems.push(blankAdvItem());
  mountAdvTab();
  toast("โหลดข้อมูลเรียบร้อย");
}

async function saveAdvForm() {
  const f = state.advForm;
  if (!f.employee_name.trim()) { toast("กรุณากรอกชื่อ-นามสกุล", true); return; }
  const payload = { ...f };
  delete payload.created_at; delete payload.updated_at; delete payload.created_by;
  payload.transfer_to_company_date = payload.transfer_to_company_date || null;
  payload.company_return_to_employee_date = payload.company_return_to_employee_date || null;
  let formId = f.id;

  try {
    if (formId) {
      const { error } = await sb.from("adv_clear_forms").update(payload).eq("id", formId);
      if (error) throw error;
    } else {
      const { data: userData } = await sb.auth.getUser();
      payload.created_by = userData.user.id;
      delete payload.id;
      const { data, error } = await sb.from("adv_clear_forms").insert(payload).select().single();
      if (error) throw error;
      formId = data.id;
      state.advForm.id = formId;
    }

    // เขียนรายการใหม่ทั้งหมด (ลบของเดิมแล้วเพิ่มใหม่ เพื่อความง่ายและถูกต้องเสมอ)
    await sb.from("adv_clear_items").delete().eq("form_id", formId);
    const itemPayload = state.advItems.map((it, idx) => {
      const row = {
        form_id: formId, sort_order: idx, item_month: it.item_month, item_date: it.item_date || null,
        description: it.description, clear_amount: Number(it.clear_amount) || 0, has_bill: !!it.has_bill,
        no_bill: !it.has_bill,
      };
      ADV_COST_FIELDS.forEach(fdef => row[fdef.key] = 0);
      row[it._type || "cost_other"] = Number(it.clear_amount) || 0;
      return row;
    });
    const { error: insErr } = await sb.from("adv_clear_items").insert(itemPayload);
    if (insErr) throw insErr;

    toast("บันทึกข้อมูลเรียบร้อยแล้ว");
    loadAdvRecords();
  } catch (err) {
    toast("บันทึกไม่สำเร็จ: " + err.message, true);
  }
}

async function deleteAdvForm() {
  if (!state.advForm.id) return;
  if (!confirm("ยืนยันลบใบนี้ทั้งหมด? การลบไม่สามารถย้อนกลับได้")) return;
  const { error } = await sb.from("adv_clear_forms").delete().eq("id", state.advForm.id);
  if (error) { toast("ลบไม่สำเร็จ: " + error.message, true); return; }
  toast("ลบเรียบร้อย");
  newAdvForm();
  loadAdvRecords();
}

/* ============================================================
   FM-AC-06: ใบสรุปค่าใช้จ่าย (ไม่มีเอกสาร)
   ============================================================ */
function newExpForm(doRender = true) {
  state.expForm = blankExpForm();
  state.expItems = [blankExpItem()];
  if (doRender) mountExpTab();
}

function expTotal() {
  return state.expItems.reduce((s, it) => s + (Number(it.amount) || 0), 0);
}

function mountExpTab() {
  const target = $("#tab-content");
  const f = state.expForm;
  const total = expTotal();

  target.innerHTML = `
    <div class="grid-2">
      <div>
        <div class="card">
          <h2><span class="eyebrow">ข้อมูลผู้เบิก</span>ใบสรุปค่าใช้จ่าย (ไม่มีเอกสาร)</h2>
          <div class="field-grid">
            <div class="field"><label>ชื่อ-นามสกุล</label><input id="e-employee_name" value="${esc(f.employee_name)}" /></div>
            <div class="field"><label>ตำแหน่ง</label><input id="e-position" value="${esc(f.position)}" /></div>
            <div class="field"><label>แผนก</label><input id="e-department" value="${esc(f.department)}" /></div>
            <div class="field span-2"><label>โครงการ</label><input id="e-project" value="${esc(f.project)}" /></div>
          </div>
        </div>

        <div class="card">
          <h2><span class="eyebrow">รายการค่าใช้จ่าย</span>รายการ (ไม่มีเอกสาร/บิล)</h2>
          <div class="table-scroll">
            <table class="items">
              <thead>
                <tr>
                  <th>ลำดับ</th><th>วันที่</th><th style="min-width:200px">รายการ</th><th>จำนวนเงิน</th>
                  <th>เริ่มต้น</th><th>สิ้นสุด</th><th>จำนวนคน</th><th style="min-width:160px">รายชื่อผู้ร่วมเดินทาง</th>
                  <th>อื่นๆ</th><th style="min-width:120px">หมายเหตุ</th><th></th>
                </tr>
              </thead>
              <tbody>
                ${state.expItems.map((it, idx) => expItemRow(it, idx)).join("")}
              </tbody>
            </table>
          </div>
          <button class="add-row-btn no-print" id="exp-add-row">+ เพิ่มรายการ</button>
        </div>

        <div class="card no-print">
          <h2><span class="eyebrow">ผู้เกี่ยวข้อง</span>ผู้ขอเบิก / ผู้อนุมัติ</h2>
          <div class="field-grid">
            <div class="field"><label>ผู้ขอเบิก</label><input id="e-requester_name" value="${esc(f.requester_name)}" /></div>
            <div class="field"><label>ผู้อนุมัติ</label><input id="e-approver_name" value="${esc(f.approver_name)}" /></div>
            <div class="field"><label>สถานะ</label>
              <select id="e-status">
                ${["draft","submitted","approved"].map(s => `<option value="${s}" ${f.status===s?"selected":""}>${statusLabel(s)}</option>`).join("")}
              </select>
            </div>
          </div>
        </div>
      </div>

      <div>
        <div class="ledger">
          <div class="ledger-head">
            <span class="eyebrow">สรุปยอด</span>
            <h3>ใบสรุปค่าใช้จ่าย (ไม่มีเอกสาร)</h3>
          </div>
          <div class="ledger-perf"></div>
          <div class="ledger-body" id="exp-ledger-body">
            <div class="ledger-line"><span>รวมเงิน</span><span class="amt">฿${fmt(total)}</span></div>
            <div class="helper-text">(${esc(thaiBahtText(total))})</div>
          </div>
        </div>

        <div class="card no-print" style="margin-top:18px">
          <h2>บันทึกข้อมูล</h2>
          <div class="action-row">
            <button class="btn btn-primary" id="exp-save-btn">บันทึกลง Supabase</button>
            <button class="btn btn-ghost" id="exp-new-btn">สร้างใบใหม่</button>
            <button class="btn btn-ghost" id="exp-print-btn">พิมพ์ / PDF</button>
            ${f.id ? `<button class="btn btn-danger" id="exp-delete-btn">ลบใบนี้</button>` : ""}
          </div>
        </div>

        <div class="card no-print" style="margin-top:18px">
          <h2>รายการที่บันทึกไว้</h2>
          <div class="list-toolbar">
            <input id="exp-search" placeholder="ค้นหาชื่อ / โครงการ" value="${esc(state.expSearch)}" />
          </div>
          <div id="exp-records-list">${renderExpRecordsTable()}</div>
        </div>
      </div>
    </div>
  `;

  bindExpEvents();
  loadExpRecords();
}

function expItemRow(it, idx) {
  return `
  <tr>
    <td><input class="e-item-field" data-field="seq" data-idx="${idx}" type="number" value="${it.seq}" style="width:50px" /></td>
    <td><input class="e-item-field" data-field="item_date" data-idx="${idx}" type="date" value="${it.item_date || ""}" /></td>
    <td><input class="e-item-field" data-field="description" data-idx="${idx}" value="${esc(it.description)}" /></td>
    <td><input class="e-item-field" data-field="amount" data-idx="${idx}" type="number" step="0.01" value="${it.amount}" /></td>
    <td><input class="e-item-field" data-field="travel_start" data-idx="${idx}" value="${esc(it.travel_start)}" style="width:80px" /></td>
    <td><input class="e-item-field" data-field="travel_end" data-idx="${idx}" value="${esc(it.travel_end)}" style="width:80px" /></td>
    <td><input class="e-item-field" data-field="guest_count" data-idx="${idx}" type="number" value="${it.guest_count}" style="width:60px" /></td>
    <td><input class="e-item-field" data-field="guest_names" data-idx="${idx}" value="${esc(it.guest_names)}" /></td>
    <td><input class="e-item-field" data-field="other_note" data-idx="${idx}" value="${esc(it.other_note)}" /></td>
    <td><input class="e-item-field" data-field="remark" data-idx="${idx}" value="${esc(it.remark)}" /></td>
    <td class="center"><button class="row-del e-item-del" data-idx="${idx}">✕</button></td>
  </tr>`;
}

function bindExpEvents() {
  $("#exp-add-row").addEventListener("click", () => {
    const nextSeq = state.expItems.length + 1;
    const item = blankExpItem();
    item.seq = nextSeq;
    state.expItems.push(item);
    mountExpTab();
  });
  $$(".e-item-field").forEach(el => {
    el.addEventListener("input", () => {
      const idx = Number(el.dataset.idx);
      state.expItems[idx][el.dataset.field] = el.value;
      if (el.dataset.field === "amount") refreshExpSummaryOnly();
    });
  });
  $$(".e-item-del").forEach(btn => {
    btn.addEventListener("click", () => {
      const idx = Number(btn.dataset.idx);
      state.expItems.splice(idx, 1);
      if (state.expItems.length === 0) state.expItems.push(blankExpItem());
      mountExpTab();
    });
  });

  ["employee_name","position","department","project","requester_name","approver_name","status"].forEach(key => {
    const el = $(`#e-${key}`);
    if (!el) return;
    el.addEventListener("input", () => state.expForm[key] = el.value);
  });

  $("#exp-save-btn").addEventListener("click", saveExpForm);
  $("#exp-new-btn").addEventListener("click", () => newExpForm());
  $("#exp-print-btn").addEventListener("click", () => window.print());
  const delBtn = $("#exp-delete-btn");
  if (delBtn) delBtn.addEventListener("click", deleteExpForm);
  $("#exp-search").addEventListener("input", (e) => {
    state.expSearch = e.target.value;
    $("#exp-records-list").innerHTML = renderExpRecordsTable();
    bindExpRecordClicks();
  });
  bindExpRecordClicks();
}

function refreshExpSummaryOnly() {
  const total = expTotal();
  $("#exp-ledger-body").innerHTML = `
    <div class="ledger-line"><span>รวมเงิน</span><span class="amt">฿${fmt(total)}</span></div>
    <div class="helper-text">(${esc(thaiBahtText(total))})</div>`;
}

function renderExpRecordsTable() {
  const q = state.expSearch.trim().toLowerCase();
  const rows = state.expRecords.filter(r =>
    !q || (r.employee_name || "").toLowerCase().includes(q) || (r.project || "").toLowerCase().includes(q));
  if (rows.length === 0) return `<div class="empty-state">ยังไม่มีรายการที่บันทึกไว้ — กรอกฟอร์มด้านบนแล้วกดบันทึก</div>`;
  return `
  <table class="records">
    <thead><tr><th>ชื่อ</th><th>โครงการ</th><th>สถานะ</th></tr></thead>
    <tbody>
      ${rows.map(r => `<tr data-id="${r.id}"><td>${esc(r.employee_name)}</td><td>${esc(r.project || "-")}</td><td><span class="status-tag ${r.status}">${statusLabel(r.status)}</span></td></tr>`).join("")}
    </tbody>
  </table>`;
}

function bindExpRecordClicks() {
  $$("#exp-records-list tr[data-id]").forEach(tr => tr.addEventListener("click", () => loadExpForm(tr.dataset.id)));
}

async function loadExpRecords() {
  const { data, error } = await sb.from("exp_noreceipt_forms").select("*").order("created_at", { ascending: false });
  if (error) { toast("โหลดรายการไม่สำเร็จ: " + error.message, true); return; }
  state.expRecords = data || [];
  const listEl = $("#exp-records-list");
  if (listEl) { listEl.innerHTML = renderExpRecordsTable(); bindExpRecordClicks(); }
}

async function loadExpForm(id) {
  const { data: form, error } = await sb.from("exp_noreceipt_forms").select("*").eq("id", id).single();
  if (error) { toast("โหลดข้อมูลไม่สำเร็จ: " + error.message, true); return; }
  const { data: items, error: itemErr } = await sb.from("exp_noreceipt_items").select("*").eq("form_id", id).order("seq");
  if (itemErr) { toast("โหลดรายการไม่สำเร็จ: " + itemErr.message, true); return; }
  state.expForm = form;
  state.expItems = (items || []).map(it => ({ ...it, _uid: uid() }));
  if (state.expItems.length === 0) state.expItems.push(blankExpItem());
  mountExpTab();
  toast("โหลดข้อมูลเรียบร้อย");
}

async function saveExpForm() {
  const f = state.expForm;
  if (!f.employee_name.trim()) { toast("กรุณากรอกชื่อ-นามสกุล", true); return; }
  const payload = { ...f };
  delete payload.created_at; delete payload.updated_at; delete payload.created_by;
  let formId = f.id;

  try {
    if (formId) {
      const { error } = await sb.from("exp_noreceipt_forms").update(payload).eq("id", formId);
      if (error) throw error;
    } else {
      const { data: userData } = await sb.auth.getUser();
      payload.created_by = userData.user.id;
      delete payload.id;
      const { data, error } = await sb.from("exp_noreceipt_forms").insert(payload).select().single();
      if (error) throw error;
      formId = data.id;
      state.expForm.id = formId;
    }

    await sb.from("exp_noreceipt_items").delete().eq("form_id", formId);
    const itemPayload = state.expItems.map(it => ({
      form_id: formId, seq: Number(it.seq) || 0, item_date: it.item_date || null, description: it.description,
      amount: Number(it.amount) || 0, travel_start: it.travel_start, travel_end: it.travel_end,
      guest_count: it.guest_count ? Number(it.guest_count) : null, guest_names: it.guest_names,
      other_note: it.other_note, remark: it.remark,
    }));
    const { error: insErr } = await sb.from("exp_noreceipt_items").insert(itemPayload);
    if (insErr) throw insErr;

    toast("บันทึกข้อมูลเรียบร้อยแล้ว");
    loadExpRecords();
  } catch (err) {
    toast("บันทึกไม่สำเร็จ: " + err.message, true);
  }
}

async function deleteExpForm() {
  if (!state.expForm.id) return;
  if (!confirm("ยืนยันลบใบนี้ทั้งหมด? การลบไม่สามารถย้อนกลับได้")) return;
  const { error } = await sb.from("exp_noreceipt_forms").delete().eq("id", state.expForm.id);
  if (error) { toast("ลบไม่สำเร็จ: " + error.message, true); return; }
  toast("ลบเรียบร้อย");
  newExpForm();
  loadExpRecords();
}

/* ============================================================
   Shared helpers
   ============================================================ */
function statusLabel(s) {
  return { draft: "ฉบับร่าง", submitted: "ส่งอนุมัติแล้ว", approved: "อนุมัติแล้ว" }[s] || s;
}
function esc(str) {
  if (str === null || str === undefined) return "";
  return String(str).replace(/[&<>"']/g, c => ({ "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;" }[c]));
}

/* ============================================================
   Boot
   ============================================================ */
initAuth();
