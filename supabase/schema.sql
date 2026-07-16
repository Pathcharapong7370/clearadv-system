-- ============================================================
-- ระบบเคลียร์เงินทดรองจ่าย (FM-AC-05) และ ใบสรุปค่าใช้จ่ายไม่มีเอกสาร (FM-AC-06)
-- Schema สำหรับ Supabase (Postgres)
-- วิธีใช้: เปิด Supabase Dashboard -> SQL Editor -> วางไฟล์นี้ทั้งหมด -> กด Run
-- ============================================================

create extension if not exists "pgcrypto";

-- ------------------------------------------------------------
-- FM-AC-05: ใบสรุปเคลียร์เงินทดรองจ่าย (หัวเอกสาร)
-- ------------------------------------------------------------
create table if not exists adv_clear_forms (
  id                              uuid primary key default gen_random_uuid(),
  created_at                      timestamptz not null default now(),
  updated_at                      timestamptz not null default now(),
  created_by                      uuid references auth.users(id),
  employee_name                   text not null,
  department                      text,
  project                         text,
  province                        text,
  doc_year                        int,
  adv_amount                      numeric(12,2) default 0,   -- จำนวนเงิน Advance
  adv_ref                         text,                        -- อ้างอิง ADV
  transfer_to_company_date        date,   -- โอนเงินคืนบริษัทวันที่
  company_return_to_employee_date date,   -- บริษัทฯ คืนเงินพนักงานวันที่
  transfer_proof_note             text,   -- หลักฐานการโอนเงิน (เช่น เลขที่ slip)
  requester_name                  text,   -- ผู้ขอเบิก
  approver_name                   text,   -- ผู้อนุมัติ
  receiver_name                   text,   -- ผู้รับเงิน
  checker_name                    text,   -- ผู้ตรวจสอบ
  accountant_name                 text,   -- ผู้บันทึกบัญชี
  status                          text not null default 'draft'
                                    check (status in ('draft','submitted','approved'))
);

-- รายการค่าใช้จ่ายของแต่ละใบ (FM-AC-05)
create table if not exists adv_clear_items (
  id                uuid primary key default gen_random_uuid(),
  form_id           uuid not null references adv_clear_forms(id) on delete cascade,
  sort_order        int not null default 0,
  item_month        text,        -- เดือน
  item_date         date,        -- วันที่
  description       text,        -- รายการ
  clear_amount      numeric(12,2) default 0,  -- เคลียร์คชจ. จำนวนเงิน (ยอด Advance รวมอยู่ที่หัวเอกสาร adv_clear_forms.adv_amount)
  has_bill          boolean default false,    -- มีบิล
  no_bill           boolean default false,    -- ไม่มีบิล
  cost_phone        numeric(12,2) default 0,  -- ค่าโทรศัพท์
  cost_travel       numeric(12,2) default 0,  -- ค่าเดินทาง
  cost_postage      numeric(12,2) default 0,  -- ค่าไปรษณีย์
  cost_taxi         numeric(12,2) default 0,  -- ค่าแท็กซี่
  cost_fuel         numeric(12,2) default 0,  -- ค่าน้ำมัน
  cost_tollway      numeric(12,2) default 0,  -- ค่าทางด่วน
  cost_food         numeric(12,2) default 0,  -- ค่าอาหาร
  cost_hotel        numeric(12,2) default 0,  -- ค่าที่พัก
  cost_stationery   numeric(12,2) default 0,  -- เครื่องเขียนแบบพิมพ์
  cost_other        numeric(12,2) default 0   -- อื่นๆ
);

create index if not exists idx_adv_clear_items_form_id on adv_clear_items(form_id);

-- ------------------------------------------------------------
-- FM-AC-06: ใบสรุปค่าใช้จ่าย (ไม่มีเอกสาร) (หัวเอกสาร)
-- ------------------------------------------------------------
create table if not exists exp_noreceipt_forms (
  id              uuid primary key default gen_random_uuid(),
  created_at      timestamptz not null default now(),
  updated_at      timestamptz not null default now(),
  created_by      uuid references auth.users(id),
  employee_name   text not null,
  position        text,   -- ตำแหน่ง
  department      text,
  project         text,
  requester_name  text,   -- ผู้ขอเบิก
  approver_name   text,   -- ผู้อนุมัติ
  status          text not null default 'draft'
                    check (status in ('draft','submitted','approved'))
);

create table if not exists exp_noreceipt_items (
  id            uuid primary key default gen_random_uuid(),
  form_id       uuid not null references exp_noreceipt_forms(id) on delete cascade,
  seq           int,
  item_date     date,
  description   text,
  amount        numeric(12,2) default 0,
  travel_start  text,   -- ค่าเดินทาง: เริ่มต้น
  travel_end    text,   -- ค่าเดินทาง: สิ้นสุด
  guest_count   int,    -- ค่ารับรอง: จำนวน (คน)
  guest_names   text,   -- ค่ารับรอง: รายชื่อ
  other_note    text,   -- อื่นๆ
  remark        text    -- หมายเหตุ
);

create index if not exists idx_exp_noreceipt_items_form_id on exp_noreceipt_items(form_id);

-- ------------------------------------------------------------
-- ฟังก์ชัน + trigger อัปเดต updated_at อัตโนมัติ
-- ------------------------------------------------------------
create or replace function set_updated_at()
returns trigger as $$
begin
  new.updated_at = now();
  return new;
end;
$$ language plpgsql;

drop trigger if exists trg_adv_clear_forms_updated on adv_clear_forms;
create trigger trg_adv_clear_forms_updated
  before update on adv_clear_forms
  for each row execute function set_updated_at();

drop trigger if exists trg_exp_noreceipt_forms_updated on exp_noreceipt_forms;
create trigger trg_exp_noreceipt_forms_updated
  before update on exp_noreceipt_forms
  for each row execute function set_updated_at();

-- ------------------------------------------------------------
-- Row Level Security: อนุญาตเฉพาะผู้ใช้ที่ล็อกอิน (authenticated) เท่านั้น
-- ------------------------------------------------------------
alter table adv_clear_forms       enable row level security;
alter table adv_clear_items       enable row level security;
alter table exp_noreceipt_forms   enable row level security;
alter table exp_noreceipt_items   enable row level security;

drop policy if exists "authenticated full access" on adv_clear_forms;
create policy "authenticated full access" on adv_clear_forms
  for all using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');

drop policy if exists "authenticated full access" on adv_clear_items;
create policy "authenticated full access" on adv_clear_items
  for all using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');

drop policy if exists "authenticated full access" on exp_noreceipt_forms;
create policy "authenticated full access" on exp_noreceipt_forms
  for all using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');

drop policy if exists "authenticated full access" on exp_noreceipt_items;
create policy "authenticated full access" on exp_noreceipt_items
  for all using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');
