import { createClient } from "https://esm.sh/@supabase/supabase-js@2";
import * as XLSX from "https://esm.sh/xlsx@0.18.5";
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType, BorderStyle, VerticalAlign } from "https://esm.sh/docx@8.5.0";

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

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

function thaiDate(iso) {
  if (!iso) return "";
  const months = ["ม.ค.","ก.พ.","มี.ค.","เม.ย.","พ.ค.","มิ.ย.","ก.ค.","ส.ค.","ก.ย.","ต.ค.","พ.ย.","ธ.ค."];
  const d = new Date(iso);
  if (isNaN(d)) return iso;
  return `${d.getDate()} ${months[d.getMonth()]} ${d.getFullYear() + 543}`;
}

/* ============================================================
   State
   ============================================================ */
const state = {
  tab: "adv",           // 'adv' = FM-AC-05, 'exp' = FM-AC-06, 'proj' = โครงการ
  session: null,
  advForm: null,        // header object being edited (null id = new)
  advItems: [],
  advRecords: [],
  advSearch: "",
  expForm: null,
  expItems: [],
  expRecords: [],
  expSearch: "",
  projSubTab: "info",   // 'info' = ข้อมูลโครงการ, 'letter' = แบบฟอร์มแจ้งทำหนังสือออก
  projectForm: null,
  scheduleItems: [],
  projectMembers: [],
  projects: [],
  projSearch: "",
  letterForm: null,
  letterPeriods: [],
  letterAttendees: [],
  letterRecords: [],
  letterSearch: "",
};

const SYSTEM_TYPE_OPTIONS = ["BMS-HOSxP", "BMS-HOSxP XE", "BMS Data Center", "BMS-INVENTORY", "อื่นๆ"];
const ACTION_TYPE_OPTIONS = ["เข้าปฏิบัติงานติดตั้ง", "Re-visit", "MA", "สำรวจระบบ", "ขอคัดลอกฐานข้อมูล", "นำเสนอโปรแกรม", "ตอบกลับวิทยากร", "ส่งมอบงาน", "อื่นๆ"];
const RECIPIENT_TYPE_OPTIONS = ["ผู้อำนวยการโรงพยาบาล", "นายแพทย์สำนักงานสาธารณสุข", "อื่นๆ"];
const PURPOSE_TYPE_OPTIONS = ["เพื่อทราบ", "ให้คณะกรรมการตรวจรับและเบิกจ่ายเงิน"];

function blankProjectForm() {
  return {
    id: null, project_name: "", customer_name: "", province: "", start_date: todayISO(), end_date: todayISO(),
    objectives: "", notes: "", preparer_name: "", preparer_position: "", reviewer_name: "", reviewer_position: "",
    status: "active",
  };
}
function blankScheduleItem() {
  return { _uid: uid(), work_date_label: "", time_range: "", task_detail: "", bms_responsible: "", hospital_responsible: "", hospital_preparation: "" };
}
function blankProjectMember() {
  return { _uid: uid(), member_name: "", member_position: "", member_phone: "" };
}
function blankLetterForm() {
  return {
    id: null, project_id: "", request_date: todayISO(), requester_name: "", requester_position: "", requester_department: "",
    system_types: [], system_other_note: "", action_types: [], action_revisit_no: "", action_ma_no: "", action_delivery_no: "",
    action_other_note: "", attachment_count: 1, attachment_pages: 1, recipient_type: RECIPIENT_TYPE_OPTIONS[0], recipient_other: "",
    purpose_type: PURPOSE_TYPE_OPTIONS[0], contract_no: "", contract_amount: 0, contract_date: "", quotation_no: "",
    delivery_email: "", delivery_contact_name: "", delivery_contact_phone: "", cc_email: "", travel_method: "",
  };
}
function blankLetterPeriod() {
  return { _uid: uid(), date_range_text: "", location: "" };
}
function blankLetterAttendee() {
  return { _uid: uid(), attendee_name: "", attendee_position: "" };
}

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
    id: null, employee_name: "", department: "", project: "", project_id: "", province: "",
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
  return { id: null, employee_name: "", position: "", department: "", project: "", project_id: "", requester_name: "", approver_name: "", status: "draft" };
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
  if (state.session) await loadProjectsList();
  sb.auth.onAuthStateChange(async (_event, session) => {
    state.session = session;
    if (session) await loadProjectsList();
    renderRoot();
  });
  renderRoot();
}

async function loadProjectsList() {
  const { data, error } = await sb.from("projects").select("id, project_name, customer_name").order("created_at", { ascending: false });
  if (!error) state.projects = data || [];
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
  } else if (state.tab === "exp") {
    if (!state.expForm) newExpForm(false);
    mountExpTab();
  } else {
    if (!state.projectForm) newProjectForm(false);
    if (!state.letterForm) newLetterForm(false);
    mountProjTab();
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
    <button class="tab-btn ${state.tab === "proj" ? "active" : ""}" data-tab="proj">โครงการ &amp; แจ้งทำหนังสือออก</button>
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
            <div class="field span-2">
              <label>โครงการ</label>
              <select id="a-project_id">
                <option value="">-- ไม่ผูกกับโครงการ / พิมพ์เอง --</option>
                ${state.projects.map(p => `<option value="${p.id}" ${f.project_id===p.id?"selected":""}>${esc(p.project_name)}${p.customer_name?` — ${esc(p.customer_name)}`:""}</option>`).join("")}
              </select>
              <input id="a-project" value="${esc(f.project)}" placeholder="ชื่อโครงการ (เติมอัตโนมัติเมื่อเลือกด้านบน หรือพิมพ์เองได้)" style="margin-top:6px" />
            </div>
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
  const advProjectSel = $("#a-project_id");
  if (advProjectSel) {
    advProjectSel.addEventListener("change", () => {
      const pid = advProjectSel.value;
      state.advForm.project_id = pid || null;
      const proj = state.projects.find(p => p.id === pid);
      if (proj) state.advForm.project = proj.project_name;
      mountAdvTab();
    });
  }

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
            <div class="field span-2">
              <label>โครงการ</label>
              <select id="e-project_id">
                <option value="">-- ไม่ผูกกับโครงการ / พิมพ์เอง --</option>
                ${state.projects.map(p => `<option value="${p.id}" ${f.project_id===p.id?"selected":""}>${esc(p.project_name)}${p.customer_name?` — ${esc(p.customer_name)}`:""}</option>`).join("")}
              </select>
              <input id="e-project" value="${esc(f.project)}" placeholder="ชื่อโครงการ (เติมอัตโนมัติเมื่อเลือกด้านบน หรือพิมพ์เองได้)" style="margin-top:6px" />
            </div>
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
  const expProjectSel = $("#e-project_id");
  if (expProjectSel) {
    expProjectSel.addEventListener("change", () => {
      const pid = expProjectSel.value;
      state.expForm.project_id = pid || null;
      const proj = state.projects.find(p => p.id === pid);
      if (proj) state.expForm.project = proj.project_name;
      mountExpTab();
    });
  }

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
   โครงการ: ข้อมูลโครงการ + ตารางดำเนินงาน + ทีมงาน
   + แบบฟอร์มแจ้งทำหนังสือออก (ผูกกับโครงการ)
   ============================================================ */
function newProjectForm(doRender = true) {
  state.projectForm = blankProjectForm();
  state.scheduleItems = [blankScheduleItem()];
  state.projectMembers = [blankProjectMember()];
  if (doRender) mountProjTab();
}
function newLetterForm(doRender = true) {
  state.letterForm = blankLetterForm();
  state.letterPeriods = [blankLetterPeriod()];
  state.letterAttendees = [blankLetterAttendee()];
  if (doRender) mountProjTab();
}

function mountProjTab() {
  const target = $("#tab-content");
  target.innerHTML = `
    <div class="tabbar" style="margin:0 0 18px;padding:0">
      <button class="tab-btn ${state.projSubTab === "info" ? "active" : ""}" data-subtab="info">ข้อมูลโครงการ &amp; ตารางดำเนินงาน</button>
      <button class="tab-btn ${state.projSubTab === "letter" ? "active" : ""}" data-subtab="letter">แบบฟอร์มแจ้งทำหนังสือออก</button>
    </div>
    <div id="proj-sub-content"></div>
  `;
  $$(".tabbar button[data-subtab]").forEach(btn => btn.addEventListener("click", () => {
    state.projSubTab = btn.dataset.subtab;
    mountProjTab();
  }));
  if (state.projSubTab === "info") mountProjectInfoSection();
  else mountLetterSection();
}

/* ---------------- ข้อมูลโครงการ ---------------- */
function mountProjectInfoSection() {
  const sub = $("#proj-sub-content");
  const f = state.projectForm;

  sub.innerHTML = `
    <div class="grid-2">
      <div>
        <div class="card">
          <h2><span class="eyebrow">ข้อมูลหลัก</span>ข้อมูลโครงการ</h2>
          <div class="field-grid">
            <div class="field span-2"><label>ชื่อโครงการ</label><input id="p-project_name" value="${esc(f.project_name)}" placeholder="เช่น ติดตั้งระบบครุภัณฑ์ โรงพยาบาลสันติสุข" /></div>
            <div class="field"><label>ชื่อหน่วยงาน/โรงพยาบาล</label><input id="p-customer_name" value="${esc(f.customer_name)}" /></div>
            <div class="field"><label>จังหวัด</label><input id="p-province" value="${esc(f.province)}" /></div>
            <div class="field"><label>วันที่เริ่ม MasterPlan</label><input id="p-start_date" type="date" value="${f.start_date || ""}" /></div>
            <div class="field"><label>วันที่สิ้นสุด MasterPlan</label><input id="p-end_date" type="date" value="${f.end_date || ""}" /></div>
            <div class="field span-2"><label>วัตถุประสงค์ (1 บรรทัดต่อ 1 ข้อ)</label><textarea id="p-objectives" rows="3">${esc(f.objectives)}</textarea></div>
            <div class="field span-2"><label>หมายเหตุ (1 บรรทัดต่อ 1 ข้อ)</label><textarea id="p-notes" rows="3">${esc(f.notes)}</textarea></div>
            <div class="field"><label>ผู้จัดทำ</label><input id="p-preparer_name" value="${esc(f.preparer_name)}" /></div>
            <div class="field"><label>ตำแหน่งผู้จัดทำ</label><input id="p-preparer_position" value="${esc(f.preparer_position)}" /></div>
            <div class="field"><label>ผู้ตรวจสอบ</label><input id="p-reviewer_name" value="${esc(f.reviewer_name)}" /></div>
            <div class="field"><label>ตำแหน่งผู้ตรวจสอบ</label><input id="p-reviewer_position" value="${esc(f.reviewer_position)}" /></div>
            <div class="field"><label>สถานะโครงการ</label>
              <select id="p-status">
                ${[["active","ดำเนินการอยู่"],["completed","เสร็จสิ้น"],["cancelled","ยกเลิก"]].map(([v,l]) => `<option value="${v}" ${f.status===v?"selected":""}>${l}</option>`).join("")}
              </select>
            </div>
          </div>
        </div>

        <div class="card">
          <h2><span class="eyebrow">MasterPlan</span>ตารางการดำเนินงาน</h2>
          <div class="table-scroll">
            <table class="items" id="sched-table">
              <thead><tr><th style="min-width:120px">วันที่</th><th style="min-width:110px">เวลา</th><th style="min-width:220px">รายละเอียดงาน</th><th style="min-width:120px">ผู้รับผิดชอบ BMS</th><th style="min-width:120px">ผู้รับผิดชอบโรงพยาบาล</th><th style="min-width:160px">การเตรียมตัวของ รพ.</th><th></th></tr></thead>
              <tbody>${state.scheduleItems.map((it, idx) => scheduleItemRow(it, idx)).join("")}</tbody>
            </table>
          </div>
          <button class="add-row-btn no-print" id="sched-add-row">+ เพิ่มรายการ</button>
        </div>

        <div class="card">
          <h2><span class="eyebrow">ทีมงาน</span>ผู้ปฏิบัติงานประจำโครงการ</h2>
          <div class="table-scroll">
            <table class="items" id="member-table">
              <thead><tr><th style="min-width:180px">ชื่อ-นามสกุล</th><th style="min-width:180px">ตำแหน่ง</th><th style="min-width:130px">เบอร์โทร</th><th></th></tr></thead>
              <tbody>${state.projectMembers.map((it, idx) => memberRow(it, idx)).join("")}</tbody>
            </table>
          </div>
          <button class="add-row-btn no-print" id="member-add-row">+ เพิ่มทีมงาน</button>
        </div>
      </div>

      <div>
        <div class="card no-print">
          <h2>บันทึก &amp; ส่งออก</h2>
          <div class="action-row">
            <button class="btn btn-primary" id="proj-save-btn">บันทึกลง Supabase</button>
            <button class="btn btn-ghost" id="proj-new-btn">สร้างโครงการใหม่</button>
            <button class="btn btn-ghost" id="proj-export-btn">ส่งออกตารางดำเนินงาน (Excel)</button>
            <button class="btn btn-ghost" id="proj-print-btn">พิมพ์ / PDF</button>
            ${f.id ? `<button class="btn btn-danger" id="proj-delete-btn">ลบโครงการนี้</button>` : ""}
          </div>
          <p class="helper-text">ไฟล์ Excel ที่ส่งออกจะมี 2 ชีต: ตารางดำเนินงาน และรายชื่อผู้ปฏิบัติงาน ตามรูปแบบต้นฉบับ</p>
        </div>

        <div class="card no-print" style="margin-top:18px">
          <h2>โครงการที่บันทึกไว้</h2>
          <div class="list-toolbar"><input id="proj-search" placeholder="ค้นหาชื่อโครงการ / หน่วยงาน" value="${esc(state.projSearch)}" /></div>
          <div id="proj-records-list">${renderProjectRecordsTable()}</div>
        </div>
      </div>
    </div>
  `;

  bindProjectInfoEvents();
  loadProjectRecords();
}

function scheduleItemRow(it, idx) {
  return `
  <tr data-uid="${it._uid}">
    <td><input class="sched-field" data-field="work_date_label" data-idx="${idx}" value="${esc(it.work_date_label)}" placeholder="เช่น วันที่ 10 สิงหาคม 2569" /></td>
    <td><input class="sched-field" data-field="time_range" data-idx="${idx}" value="${esc(it.time_range)}" placeholder="09.00 - 12.00 น." /></td>
    <td><input class="sched-field" data-field="task_detail" data-idx="${idx}" value="${esc(it.task_detail)}" /></td>
    <td><input class="sched-field" data-field="bms_responsible" data-idx="${idx}" value="${esc(it.bms_responsible)}" /></td>
    <td><input class="sched-field" data-field="hospital_responsible" data-idx="${idx}" value="${esc(it.hospital_responsible)}" /></td>
    <td><input class="sched-field" data-field="hospital_preparation" data-idx="${idx}" value="${esc(it.hospital_preparation)}" /></td>
    <td class="center"><button class="row-del sched-del" data-idx="${idx}">✕</button></td>
  </tr>`;
}
function memberRow(it, idx) {
  return `
  <tr data-uid="${it._uid}">
    <td><input class="member-field" data-field="member_name" data-idx="${idx}" value="${esc(it.member_name)}" /></td>
    <td><input class="member-field" data-field="member_position" data-idx="${idx}" value="${esc(it.member_position)}" /></td>
    <td><input class="member-field" data-field="member_phone" data-idx="${idx}" value="${esc(it.member_phone)}" /></td>
    <td class="center"><button class="row-del member-del" data-idx="${idx}">✕</button></td>
  </tr>`;
}

function bindProjectInfoEvents() {
  $("#sched-add-row").addEventListener("click", () => { state.scheduleItems.push(blankScheduleItem()); mountProjectInfoSection(); });
  $$(".sched-field").forEach(el => el.addEventListener("input", () => {
    state.scheduleItems[Number(el.dataset.idx)][el.dataset.field] = el.value;
  }));
  $$(".sched-del").forEach(btn => btn.addEventListener("click", () => {
    state.scheduleItems.splice(Number(btn.dataset.idx), 1);
    if (state.scheduleItems.length === 0) state.scheduleItems.push(blankScheduleItem());
    mountProjectInfoSection();
  }));

  $("#member-add-row").addEventListener("click", () => { state.projectMembers.push(blankProjectMember()); mountProjectInfoSection(); });
  $$(".member-field").forEach(el => el.addEventListener("input", () => {
    state.projectMembers[Number(el.dataset.idx)][el.dataset.field] = el.value;
  }));
  $$(".member-del").forEach(btn => btn.addEventListener("click", () => {
    state.projectMembers.splice(Number(btn.dataset.idx), 1);
    if (state.projectMembers.length === 0) state.projectMembers.push(blankProjectMember());
    mountProjectInfoSection();
  }));

  ["project_name","customer_name","province","start_date","end_date","objectives","notes","preparer_name","preparer_position","reviewer_name","reviewer_position","status"].forEach(key => {
    const el = $(`#p-${key}`);
    if (!el) return;
    el.addEventListener("input", () => state.projectForm[key] = el.value);
  });

  $("#proj-save-btn").addEventListener("click", saveProjectForm);
  $("#proj-new-btn").addEventListener("click", () => newProjectForm());
  $("#proj-export-btn").addEventListener("click", exportScheduleXlsx);
  $("#proj-print-btn").addEventListener("click", () => window.print());
  const delBtn = $("#proj-delete-btn");
  if (delBtn) delBtn.addEventListener("click", deleteProjectForm);
  $("#proj-search").addEventListener("input", (e) => {
    state.projSearch = e.target.value;
    $("#proj-records-list").innerHTML = renderProjectRecordsTable();
    bindProjectRecordClicks();
  });
  bindProjectRecordClicks();
}

function renderProjectRecordsTable() {
  const q = state.projSearch.trim().toLowerCase();
  const rows = state.projects.filter(r => !q || (r.project_name || "").toLowerCase().includes(q) || (r.customer_name || "").toLowerCase().includes(q));
  if (rows.length === 0) return `<div class="empty-state">ยังไม่มีโครงการที่บันทึกไว้</div>`;
  return `<table class="records"><thead><tr><th>ชื่อโครงการ</th><th>หน่วยงาน</th></tr></thead><tbody>
    ${rows.map(r => `<tr data-id="${r.id}"><td>${esc(r.project_name)}</td><td>${esc(r.customer_name || "-")}</td></tr>`).join("")}
  </tbody></table>`;
}
function bindProjectRecordClicks() {
  $$("#proj-records-list tr[data-id]").forEach(tr => tr.addEventListener("click", () => loadProjectForm(tr.dataset.id)));
}

async function loadProjectRecords() {
  await loadProjectsList();
  const listEl = $("#proj-records-list");
  if (listEl) { listEl.innerHTML = renderProjectRecordsTable(); bindProjectRecordClicks(); }
}

async function loadProjectForm(id) {
  const { data: form, error } = await sb.from("projects").select("*").eq("id", id).single();
  if (error) { toast("โหลดข้อมูลไม่สำเร็จ: " + error.message, true); return; }
  const { data: sched } = await sb.from("project_schedule_items").select("*").eq("project_id", id).order("sort_order");
  const { data: members } = await sb.from("project_members").select("*").eq("project_id", id).order("sort_order");
  state.projectForm = form;
  state.scheduleItems = (sched || []).map(it => ({ ...it, _uid: uid() }));
  state.projectMembers = (members || []).map(it => ({ ...it, _uid: uid() }));
  if (state.scheduleItems.length === 0) state.scheduleItems.push(blankScheduleItem());
  if (state.projectMembers.length === 0) state.projectMembers.push(blankProjectMember());
  mountProjectInfoSection();
  toast("โหลดข้อมูลโครงการเรียบร้อย");
}

async function saveProjectForm() {
  const f = state.projectForm;
  if (!f.project_name.trim()) { toast("กรุณากรอกชื่อโครงการ", true); return; }
  const payload = { ...f };
  delete payload.created_at; delete payload.updated_at; delete payload.created_by;
  payload.start_date = payload.start_date || null;
  payload.end_date = payload.end_date || null;
  let projectId = f.id;

  try {
    if (projectId) {
      const { error } = await sb.from("projects").update(payload).eq("id", projectId);
      if (error) throw error;
    } else {
      const { data: userData } = await sb.auth.getUser();
      payload.created_by = userData.user.id;
      delete payload.id;
      const { data, error } = await sb.from("projects").insert(payload).select().single();
      if (error) throw error;
      projectId = data.id;
      state.projectForm.id = projectId;
    }

    await sb.from("project_schedule_items").delete().eq("project_id", projectId);
    const schedPayload = state.scheduleItems.map((it, idx) => ({
      project_id: projectId, sort_order: idx, work_date_label: it.work_date_label, time_range: it.time_range,
      task_detail: it.task_detail, bms_responsible: it.bms_responsible, hospital_responsible: it.hospital_responsible,
      hospital_preparation: it.hospital_preparation,
    }));
    const { error: schedErr } = await sb.from("project_schedule_items").insert(schedPayload);
    if (schedErr) throw schedErr;

    await sb.from("project_members").delete().eq("project_id", projectId);
    const memberPayload = state.projectMembers.map((it, idx) => ({
      project_id: projectId, sort_order: idx, member_name: it.member_name, member_position: it.member_position, member_phone: it.member_phone,
    }));
    const { error: memberErr } = await sb.from("project_members").insert(memberPayload);
    if (memberErr) throw memberErr;

    toast("บันทึกข้อมูลโครงการเรียบร้อยแล้ว");
    loadProjectRecords();
  } catch (err) {
    toast("บันทึกไม่สำเร็จ: " + err.message, true);
  }
}

async function deleteProjectForm() {
  if (!state.projectForm.id) return;
  if (!confirm("ยืนยันลบโครงการนี้ทั้งหมด? การลบไม่สามารถย้อนกลับได้")) return;
  const { error } = await sb.from("projects").delete().eq("id", state.projectForm.id);
  if (error) { toast("ลบไม่สำเร็จ: " + error.message, true); return; }
  toast("ลบเรียบร้อย");
  newProjectForm();
  loadProjectRecords();
}

/* ---------------- ส่งออกตารางดำเนินงาน (Excel) ---------------- */
function exportScheduleXlsx() {
  const p = state.projectForm;
  const objectives = (p.objectives || "").split("\n").filter(Boolean);
  const notes = (p.notes || "").split("\n").filter(Boolean);

  const aoa = [];
  aoa.push([`ตารางการดำเนินงานโครงการ ${p.project_name || ""} ${p.customer_name || ""}`]);
  aoa.push([`จังหวัด ${p.province || "-"}`]);
  aoa.push([`แผนการปฏิบัติงาน (MasterPlan) วันที่ ${thaiDate(p.start_date)} - ${thaiDate(p.end_date)}`]);
  aoa.push([]);
  aoa.push(["วัตถุประสงค์"]);
  if (objectives.length === 0) aoa.push(["-"]);
  objectives.forEach((o, i) => aoa.push([`${i + 1}. ${o}`]));
  aoa.push([]);
  aoa.push(["วันที่", "เวลา", "รายละเอียดงาน", "ผู้รับผิดชอบ BMS", "ผู้รับผิดชอบโรงพยาบาล", "การเตรียมตัวของ รพ."]);
  state.scheduleItems.forEach(it => aoa.push([it.work_date_label, it.time_range, it.task_detail, it.bms_responsible, it.hospital_responsible, it.hospital_preparation]));
  aoa.push([]);
  aoa.push(["หมายเหตุ"]);
  if (notes.length === 0) aoa.push(["-"]);
  notes.forEach((n, i) => aoa.push([`${i + 1}. ${n}`]));
  aoa.push([]);
  aoa.push(["ผู้จัดทำ", p.preparer_name || "", "", "ผู้ตรวจสอบ", p.reviewer_name || ""]);
  aoa.push(["", p.preparer_position || "", "", "", p.reviewer_position || ""]);

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = [{ wch: 22 }, { wch: 16 }, { wch: 40 }, { wch: 20 }, { wch: 24 }, { wch: 26 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "ตารางดำเนินงาน");

  const memberAoa = [["ลำดับ", "ชื่อ-นามสกุล", "ตำแหน่ง", "เบอร์โทร"]];
  state.projectMembers.forEach((m, i) => memberAoa.push([i + 1, m.member_name, m.member_position, m.member_phone]));
  const wsMembers = XLSX.utils.aoa_to_sheet(memberAoa);
  wsMembers["!cols"] = [{ wch: 8 }, { wch: 26 }, { wch: 26 }, { wch: 16 }];
  XLSX.utils.book_append_sheet(wb, wsMembers, "ผู้ปฏิบัติงาน");

  XLSX.writeFile(wb, `ตารางดำเนินงาน_${p.project_name || "โครงการ"}.xlsx`);
  toast("ส่งออกไฟล์ Excel เรียบร้อย");
}

/* ---------------- แบบฟอร์มแจ้งทำหนังสือออก ---------------- */
function mountLetterSection() {
  const sub = $("#proj-sub-content");
  const f = state.letterForm;

  sub.innerHTML = `
    <div class="grid-2">
      <div>
        <div class="card">
          <h2><span class="eyebrow">คำขอ</span>แบบฟอร์มแจ้งทำหนังสือออก</h2>
          <div class="field-grid">
            <div class="field"><label>วันที่ยื่นคำขอ</label><input id="l-request_date" type="date" value="${f.request_date || ""}" /></div>
            <div class="field"><label>โครงการที่เกี่ยวข้อง</label>
              <select id="l-project_id">
                <option value="">-- ไม่ผูกกับโครงการ --</option>
                ${state.projects.map(p => `<option value="${p.id}" ${f.project_id===p.id?"selected":""}>${esc(p.project_name)}</option>`).join("")}
              </select>
            </div>
            <div class="field"><label>ชื่อ-นามสกุลผู้ขอ</label><input id="l-requester_name" value="${esc(f.requester_name)}" /></div>
            <div class="field"><label>ตำแหน่ง</label><input id="l-requester_position" value="${esc(f.requester_position)}" /></div>
            <div class="field span-2"><label>แผนก/ฝ่าย</label><input id="l-requester_department" value="${esc(f.requester_department)}" /></div>
            <div class="field"><label>จำนวนสิ่งที่ส่งมาด้วย (ฉบับ)</label><input id="l-attachment_count" type="number" value="${f.attachment_count}" /></div>
            <div class="field"><label>จำนวนแผ่น</label><input id="l-attachment_pages" type="number" value="${f.attachment_pages}" /></div>
          </div>
        </div>

        <div class="card">
          <h2><span class="eyebrow">ระบบงาน</span>ระบบงานที่เกี่ยวข้อง</h2>
          <div class="field-grid">
            ${SYSTEM_TYPE_OPTIONS.map(opt => `
              <label style="display:flex;align-items:center;gap:8px;font-weight:500;font-size:14px">
                <input type="checkbox" class="l-systype" value="${esc(opt)}" ${f.system_types.includes(opt)?"checked":""} /> ${esc(opt)}
              </label>`).join("")}
          </div>
          <div class="field" style="margin-top:10px"><label>ระบุเพิ่มเติม (ถ้าเลือก "อื่นๆ")</label><input id="l-system_other_note" value="${esc(f.system_other_note)}" /></div>
        </div>

        <div class="card">
          <h2><span class="eyebrow">เรื่องที่ให้ดำเนินการ</span>ประเภทงาน</h2>
          <div class="field-grid">
            ${ACTION_TYPE_OPTIONS.map(opt => `
              <label style="display:flex;align-items:center;gap:8px;font-weight:500;font-size:14px">
                <input type="checkbox" class="l-actiontype" value="${esc(opt)}" ${f.action_types.includes(opt)?"checked":""} /> ${esc(opt)}
              </label>`).join("")}
          </div>
          <div class="field-grid" style="margin-top:10px">
            <div class="field"><label>Re-visit ครั้งที่</label><input id="l-action_revisit_no" value="${esc(f.action_revisit_no)}" /></div>
            <div class="field"><label>MA ครั้งที่</label><input id="l-action_ma_no" value="${esc(f.action_ma_no)}" /></div>
            <div class="field"><label>ส่งมอบงาน งวดที่</label><input id="l-action_delivery_no" value="${esc(f.action_delivery_no)}" /></div>
            <div class="field"><label>ระบุเพิ่มเติม (ถ้าเลือก "อื่นๆ")</label><input id="l-action_other_note" value="${esc(f.action_other_note)}" /></div>
          </div>
        </div>

        <div class="card">
          <h2><span class="eyebrow">กำหนดการ</span>วันที่ดำเนินงาน &amp; สถานที่</h2>
          <div class="table-scroll">
            <table class="items"><thead><tr><th style="min-width:200px">วันที่ดำเนินงาน</th><th style="min-width:220px">สถานที่</th><th></th></tr></thead>
              <tbody>${state.letterPeriods.map((it, idx) => letterPeriodRow(it, idx)).join("")}</tbody>
            </table>
          </div>
          <button class="add-row-btn no-print" id="period-add-row">+ เพิ่มช่วงวันที่</button>
        </div>

        <div class="card">
          <h2><span class="eyebrow">ผู้เข้าปฏิบัติงาน</span>รายชื่อผู้เข้าปฏิบัติงาน</h2>
          <div class="table-scroll">
            <table class="items"><thead><tr><th style="min-width:200px">ชื่อ-นามสกุล</th><th style="min-width:220px">ตำแหน่ง</th><th></th></tr></thead>
              <tbody>${state.letterAttendees.map((it, idx) => letterAttendeeRow(it, idx)).join("")}</tbody>
            </table>
          </div>
          <button class="add-row-btn no-print" id="attendee-add-row">+ เพิ่มผู้เข้าปฏิบัติงาน</button>
        </div>

        <div class="card">
          <h2><span class="eyebrow">ปลายทาง</span>หนังสือแจ้งถึง / วัตถุประสงค์</h2>
          <div class="field-grid">
            <div class="field"><label>เรียน</label>
              <select id="l-recipient_type">${RECIPIENT_TYPE_OPTIONS.map(o => `<option value="${o}" ${f.recipient_type===o?"selected":""}>${o}</option>`).join("")}</select>
            </div>
            <div class="field"><label>ระบุเพิ่มเติม (ถ้าเลือก "อื่นๆ")</label><input id="l-recipient_other" value="${esc(f.recipient_other)}" /></div>
            <div class="field span-2"><label>หนังสือแจ้งเพื่อ</label>
              <select id="l-purpose_type">${PURPOSE_TYPE_OPTIONS.map(o => `<option value="${o}" ${f.purpose_type===o?"selected":""}>${o}</option>`).join("")}</select>
            </div>
            <div class="field"><label>สัญญาเลขที่</label><input id="l-contract_no" value="${esc(f.contract_no)}" /></div>
            <div class="field"><label>จำนวนเงิน (บาท)</label><input id="l-contract_amount" type="number" step="0.01" value="${f.contract_amount}" /></div>
            <div class="field"><label>ลงวันที่สัญญา</label><input id="l-contract_date" type="date" value="${f.contract_date || ""}" /></div>
            <div class="field"><label>ใบเสนอราคาเลขที่</label><input id="l-quotation_no" value="${esc(f.quotation_no)}" /></div>
          </div>
        </div>

        <div class="card">
          <h2><span class="eyebrow">การจัดส่ง</span>ช่องทางที่ต้องการให้จัดส่ง</h2>
          <div class="field-grid">
            <div class="field"><label>ส่งทาง E-mail</label><input id="l-delivery_email" value="${esc(f.delivery_email)}" /></div>
            <div class="field"><label>เรียน (ชื่อผู้รับ)</label><input id="l-delivery_contact_name" value="${esc(f.delivery_contact_name)}" /></div>
            <div class="field"><label>เบอร์โทร</label><input id="l-delivery_contact_phone" value="${esc(f.delivery_contact_phone)}" /></div>
            <div class="field"><label>CC E-mail</label><input id="l-cc_email" value="${esc(f.cc_email)}" /></div>
            <div class="field span-2"><label>การเดินทาง</label><input id="l-travel_method" value="${esc(f.travel_method)}" placeholder="เช่น รถเช่า, เครื่องบิน" /></div>
          </div>
        </div>
      </div>

      <div>
        <div class="card no-print">
          <h2>บันทึก &amp; ส่งออก</h2>
          <div class="action-row">
            <button class="btn btn-primary" id="letter-save-btn">บันทึกลง Supabase</button>
            <button class="btn btn-ghost" id="letter-new-btn">สร้างคำขอใหม่</button>
            <button class="btn btn-ghost" id="letter-export-btn">ส่งออกแบบฟอร์ม (Word)</button>
            <button class="btn btn-ghost" id="letter-print-btn">พิมพ์ / PDF</button>
            ${f.id ? `<button class="btn btn-danger" id="letter-delete-btn">ลบคำขอนี้</button>` : ""}
          </div>
          <p class="helper-text">ไฟล์ Word ที่ส่งออกจะจัดหน้าให้ตรงกับแบบฟอร์ม "แจ้งทำหนังสือออก" ต้นฉบับ</p>
        </div>

        <div class="card no-print" style="margin-top:18px">
          <h2>คำขอที่บันทึกไว้</h2>
          <div class="list-toolbar"><input id="letter-search" placeholder="ค้นหาชื่อผู้ขอ" value="${esc(state.letterSearch)}" /></div>
          <div id="letter-records-list">${renderLetterRecordsTable()}</div>
        </div>
      </div>
    </div>
  `;

  bindLetterEvents();
  loadLetterRecords();
}

function letterPeriodRow(it, idx) {
  return `<tr data-uid="${it._uid}">
    <td><input class="period-field" data-field="date_range_text" data-idx="${idx}" value="${esc(it.date_range_text)}" placeholder="เช่น 06 – 10 ก.ค. พ.ศ. 2569" /></td>
    <td><input class="period-field" data-field="location" data-idx="${idx}" value="${esc(it.location)}" /></td>
    <td class="center"><button class="row-del period-del" data-idx="${idx}">✕</button></td>
  </tr>`;
}
function letterAttendeeRow(it, idx) {
  return `<tr data-uid="${it._uid}">
    <td><input class="attendee-field" data-field="attendee_name" data-idx="${idx}" value="${esc(it.attendee_name)}" /></td>
    <td><input class="attendee-field" data-field="attendee_position" data-idx="${idx}" value="${esc(it.attendee_position)}" /></td>
    <td class="center"><button class="row-del attendee-del" data-idx="${idx}">✕</button></td>
  </tr>`;
}

function bindLetterEvents() {
  $("#period-add-row").addEventListener("click", () => { state.letterPeriods.push(blankLetterPeriod()); mountLetterSection(); });
  $$(".period-field").forEach(el => el.addEventListener("input", () => { state.letterPeriods[Number(el.dataset.idx)][el.dataset.field] = el.value; }));
  $$(".period-del").forEach(btn => btn.addEventListener("click", () => {
    state.letterPeriods.splice(Number(btn.dataset.idx), 1);
    if (state.letterPeriods.length === 0) state.letterPeriods.push(blankLetterPeriod());
    mountLetterSection();
  }));

  $("#attendee-add-row").addEventListener("click", () => { state.letterAttendees.push(blankLetterAttendee()); mountLetterSection(); });
  $$(".attendee-field").forEach(el => el.addEventListener("input", () => { state.letterAttendees[Number(el.dataset.idx)][el.dataset.field] = el.value; }));
  $$(".attendee-del").forEach(btn => btn.addEventListener("click", () => {
    state.letterAttendees.splice(Number(btn.dataset.idx), 1);
    if (state.letterAttendees.length === 0) state.letterAttendees.push(blankLetterAttendee());
    mountLetterSection();
  }));

  $$(".l-systype").forEach(cb => cb.addEventListener("change", () => {
    const v = cb.value;
    const arr = state.letterForm.system_types;
    const i = arr.indexOf(v);
    if (cb.checked && i === -1) arr.push(v);
    if (!cb.checked && i !== -1) arr.splice(i, 1);
  }));
  $$(".l-actiontype").forEach(cb => cb.addEventListener("change", () => {
    const v = cb.value;
    const arr = state.letterForm.action_types;
    const i = arr.indexOf(v);
    if (cb.checked && i === -1) arr.push(v);
    if (!cb.checked && i !== -1) arr.splice(i, 1);
  }));

  const letterFieldMap = ["request_date","requester_name","requester_position","requester_department","attachment_count","attachment_pages",
    "system_other_note","action_revisit_no","action_ma_no","action_delivery_no","action_other_note",
    "recipient_type","recipient_other","purpose_type","contract_no","contract_amount","contract_date","quotation_no",
    "delivery_email","delivery_contact_name","delivery_contact_phone","cc_email","travel_method"];
  const letterNumericFields = new Set(["attachment_count","attachment_pages","contract_amount"]);
  letterFieldMap.forEach(key => {
    const el = $(`#l-${key}`);
    if (!el) return;
    const evt = el.tagName === "SELECT" ? "change" : "input";
    el.addEventListener(evt, () => {
      state.letterForm[key] = letterNumericFields.has(key) ? (Number(el.value) || 0) : el.value;
    });
  });
  const letterProjectSel = $("#l-project_id");
  if (letterProjectSel) letterProjectSel.addEventListener("change", () => { state.letterForm.project_id = letterProjectSel.value || null; });

  $("#letter-save-btn").addEventListener("click", saveLetterForm);
  $("#letter-new-btn").addEventListener("click", () => newLetterForm());
  $("#letter-export-btn").addEventListener("click", exportLetterDocx);
  $("#letter-print-btn").addEventListener("click", () => window.print());
  const delBtn = $("#letter-delete-btn");
  if (delBtn) delBtn.addEventListener("click", deleteLetterForm);
  $("#letter-search").addEventListener("input", (e) => {
    state.letterSearch = e.target.value;
    $("#letter-records-list").innerHTML = renderLetterRecordsTable();
    bindLetterRecordClicks();
  });
  bindLetterRecordClicks();
}

function renderLetterRecordsTable() {
  const q = state.letterSearch.trim().toLowerCase();
  const rows = state.letterRecords.filter(r => !q || (r.requester_name || "").toLowerCase().includes(q));
  if (rows.length === 0) return `<div class="empty-state">ยังไม่มีคำขอที่บันทึกไว้</div>`;
  return `<table class="records"><thead><tr><th>ผู้ขอ</th><th>วันที่</th></tr></thead><tbody>
    ${rows.map(r => `<tr data-id="${r.id}"><td>${esc(r.requester_name)}</td><td>${esc(r.request_date || "-")}</td></tr>`).join("")}
  </tbody></table>`;
}
function bindLetterRecordClicks() {
  $$("#letter-records-list tr[data-id]").forEach(tr => tr.addEventListener("click", () => loadLetterForm(tr.dataset.id)));
}

async function loadLetterRecords() {
  const { data, error } = await sb.from("project_letters").select("*").order("created_at", { ascending: false });
  if (error) { toast("โหลดรายการไม่สำเร็จ: " + error.message, true); return; }
  state.letterRecords = data || [];
  const listEl = $("#letter-records-list");
  if (listEl) { listEl.innerHTML = renderLetterRecordsTable(); bindLetterRecordClicks(); }
}

async function loadLetterForm(id) {
  const { data: form, error } = await sb.from("project_letters").select("*").eq("id", id).single();
  if (error) { toast("โหลดข้อมูลไม่สำเร็จ: " + error.message, true); return; }
  const { data: periods } = await sb.from("project_letter_periods").select("*").eq("letter_id", id).order("sort_order");
  const { data: attendees } = await sb.from("project_letter_attendees").select("*").eq("letter_id", id).order("sort_order");
  state.letterForm = {
    ...form,
    system_types: form.system_types ? form.system_types.split(",").filter(Boolean) : [],
    action_types: form.action_types ? form.action_types.split(",").filter(Boolean) : [],
  };
  state.letterPeriods = (periods || []).map(it => ({ ...it, _uid: uid() }));
  state.letterAttendees = (attendees || []).map(it => ({ ...it, _uid: uid() }));
  if (state.letterPeriods.length === 0) state.letterPeriods.push(blankLetterPeriod());
  if (state.letterAttendees.length === 0) state.letterAttendees.push(blankLetterAttendee());
  mountLetterSection();
  toast("โหลดข้อมูลเรียบร้อย");
}

async function saveLetterForm() {
  const f = state.letterForm;
  if (!f.requester_name.trim()) { toast("กรุณากรอกชื่อผู้ขอ", true); return; }
  const payload = { ...f };
  delete payload.created_at; delete payload.updated_at; delete payload.created_by;
  payload.system_types = (f.system_types || []).join(",");
  payload.action_types = (f.action_types || []).join(",");
  payload.request_date = payload.request_date || null;
  payload.contract_date = payload.contract_date || null;
  payload.project_id = payload.project_id || null;
  let letterId = f.id;

  try {
    if (letterId) {
      const { error } = await sb.from("project_letters").update(payload).eq("id", letterId);
      if (error) throw error;
    } else {
      const { data: userData } = await sb.auth.getUser();
      payload.created_by = userData.user.id;
      delete payload.id;
      const { data, error } = await sb.from("project_letters").insert(payload).select().single();
      if (error) throw error;
      letterId = data.id;
      state.letterForm.id = letterId;
    }

    await sb.from("project_letter_periods").delete().eq("letter_id", letterId);
    const periodPayload = state.letterPeriods.map((it, idx) => ({ letter_id: letterId, sort_order: idx, date_range_text: it.date_range_text, location: it.location }));
    const { error: perErr } = await sb.from("project_letter_periods").insert(periodPayload);
    if (perErr) throw perErr;

    await sb.from("project_letter_attendees").delete().eq("letter_id", letterId);
    const attendeePayload = state.letterAttendees.map((it, idx) => ({ letter_id: letterId, sort_order: idx, attendee_name: it.attendee_name, attendee_position: it.attendee_position }));
    const { error: attErr } = await sb.from("project_letter_attendees").insert(attendeePayload);
    if (attErr) throw attErr;

    toast("บันทึกข้อมูลเรียบร้อยแล้ว");
    loadLetterRecords();
  } catch (err) {
    toast("บันทึกไม่สำเร็จ: " + err.message, true);
  }
}

async function deleteLetterForm() {
  if (!state.letterForm.id) return;
  if (!confirm("ยืนยันลบคำขอนี้ทั้งหมด? การลบไม่สามารถย้อนกลับได้")) return;
  const { error } = await sb.from("project_letters").delete().eq("id", state.letterForm.id);
  if (error) { toast("ลบไม่สำเร็จ: " + error.message, true); return; }
  toast("ลบเรียบร้อย");
  newLetterForm();
  loadLetterRecords();
}

/* ---------------- ส่งออกแบบฟอร์มแจ้งทำหนังสือออก (Word) ---------------- */
function chk(isChecked, label) {
  return `${isChecked ? "☑" : "☐"} ${label}`;
}

async function exportLetterDocx() {
  const f = state.letterForm;
  const proj = state.projects.find(p => p.id === f.project_id);

  const line = (text, opts = {}) => new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text, bold: !!opts.bold, size: opts.size || 22 })] });
  const kv = (label, value) => new Paragraph({
    spacing: { after: 100 },
    children: [new TextRun({ text: `${label} `, bold: true, size: 22 }), new TextRun({ text: value || "-", size: 22 })],
  });

  const periodRows = state.letterPeriods.map(p => new TableRow({ children: [
    new TableCell({ width: { size: 55, type: WidthType.PERCENTAGE }, children: [new Paragraph(p.date_range_text || "-")] }),
    new TableCell({ width: { size: 45, type: WidthType.PERCENTAGE }, children: [new Paragraph(p.location || "-")] }),
  ] }));
  const attendeeRows = state.letterAttendees.map((a, i) => new TableRow({ children: [
    new TableCell({ width: { size: 8, type: WidthType.PERCENTAGE }, children: [new Paragraph(String(i + 1))] }),
    new TableCell({ width: { size: 42, type: WidthType.PERCENTAGE }, children: [new Paragraph(a.attendee_name || "-")] }),
    new TableCell({ width: { size: 50, type: WidthType.PERCENTAGE }, children: [new Paragraph(a.attendee_position || "-")] }),
  ] }));

  const doc = new Document({
    sections: [{
      children: [
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [new TextRun({ text: "แบบฟอร์มแจ้งทำหนังสือออก", bold: true, size: 30 })] }),
        kv("วันที่:", thaiDate(f.request_date)),
        kv("โครงการ:", proj ? proj.project_name : f.project_name_snapshot || "-"),
        kv("ชื่อ-นามสกุล:", f.requester_name),
        kv("ตำแหน่ง:", f.requester_position),
        kv("แผนก/ฝ่าย:", f.requester_department),
        kv("สิ่งที่ส่งมาด้วย:", `จำนวน ${f.attachment_count || 0} ฉบับ ${f.attachment_pages || 0} แผ่น`),
        line("ระบบงานที่เกี่ยวข้อง:", { bold: true }),
        ...SYSTEM_TYPE_OPTIONS.map(o => line(chk(f.system_types.includes(o), o))),
        ...(f.system_other_note ? [line(`ระบุเพิ่มเติม: ${f.system_other_note}`)] : []),
        line("เรื่องที่ให้ดำเนินการ:", { bold: true }),
        ...ACTION_TYPE_OPTIONS.map(o => {
          let extra = "";
          if (o === "Re-visit" && f.action_revisit_no) extra = ` ครั้งที่ ${f.action_revisit_no}`;
          if (o === "MA" && f.action_ma_no) extra = ` ครั้งที่ ${f.action_ma_no}`;
          if (o === "ส่งมอบงาน" && f.action_delivery_no) extra = ` งวดที่ ${f.action_delivery_no}`;
          return line(chk(f.action_types.includes(o), o + extra));
        }),
        ...(f.action_other_note ? [line(`ระบุเพิ่มเติม: ${f.action_other_note}`)] : []),
        line("วันที่ดำเนินงาน / สถานที่:", { bold: true }),
        new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [
          new TableRow({ children: [
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "วันที่ดำเนินงาน", bold: true })] })] }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "สถานที่", bold: true })] })] }),
          ] }),
          ...periodRows,
        ] }),
        new Paragraph({ text: "", spacing: { after: 150 } }),
        line("รายชื่อผู้เข้าปฏิบัติงาน:", { bold: true }),
        new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [
          new TableRow({ children: [
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "ลำดับ", bold: true })] })] }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "ชื่อ-นามสกุล", bold: true })] })] }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "ตำแหน่ง", bold: true })] })] }),
          ] }),
          ...attendeeRows,
        ] }),
        new Paragraph({ text: "", spacing: { after: 150 } }),
        line("หนังสือแจ้งถึง:", { bold: true }),
        line(chk(true, `เรียน ${f.recipient_type}${f.recipient_type === "อื่นๆ" && f.recipient_other ? ": " + f.recipient_other : ""}`)),
        line("หนังสือแจ้งเพื่อ:", { bold: true }),
        line(chk(true, f.purpose_type)),
        ...(f.purpose_type === PURPOSE_TYPE_OPTIONS[1] ? [
          kv("สัญญาเลขที่:", f.contract_no), kv("จำนวนเงิน:", `${fmt(f.contract_amount)} บาท`),
          kv("ลงวันที่:", thaiDate(f.contract_date)), kv("ใบเสนอราคาเลขที่:", f.quotation_no),
        ] : []),
        line("ช่องทางที่ต้องการให้จัดส่ง:", { bold: true }),
        kv("E-mail:", f.delivery_email), kv("เรียน:", `${f.delivery_contact_name || "-"} โทร ${f.delivery_contact_phone || "-"}`),
        kv("CC E-mail:", f.cc_email), kv("การเดินทาง:", f.travel_method),
      ],
    }],
  });

  const blob = await Packer.toBlob(doc);
  downloadBlob(blob, `แจ้งทำหนังสือออก_${f.requester_name || "คำขอ"}.docx`);
  toast("ส่งออกไฟล์ Word เรียบร้อย");
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
