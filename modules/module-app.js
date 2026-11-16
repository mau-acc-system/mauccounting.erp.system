
import {
  validateLicense,
  readExcel,
  getSheet,
  addRecord,
  updateRecord,
  saveExcel
} from "../excel.js";

// ==========================================
// CONFIGURATION
// ==========================================

const cfg = window.MODULE_CONFIG;
const $ = id => document.getElementById(id);

const session = JSON.parse(
  sessionStorage.getItem("accsystem_session") || "null"
);

if (!session) {
  window.top.location.href = "../login.html";
  throw new Error("No active session");
}

if (!cfg) {
  throw new Error("Module configuration not found.");
}

const access = session.Access || [];
const admin = session.Role === "Administrator";

const allowed = perm =>
  admin ||
  access.some(x =>
    String(x.Module).trim() === cfg.permission &&
    (
      x[perm] === true ||
      String(x[perm]).toLowerCase() === "true" ||
      String(x[perm]) === "1"
    )
  );

// ==========================================
// PERMISSION CHECK
// ==========================================

if (!allowed("CanView")) {
  document.body.innerHTML = "<h2>Access denied</h2>";
  throw new Error("No view permission");
}

// ==========================================
// VARIABLES AND HELPERS
// ==========================================

let data = [];

const esc = v =>
  String(v ?? "").replace(/[&<>"']/g, c => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;"
  }[c]));

const n = v => Number(v) || 0;

const now = () => new Date().toISOString();

const withUpdatedAt = updates =>
  cfg.updatedAtField === false
    ? updates
    : { ...updates, UpdatedAt: now() };

const id = (prefix, arr, key) =>
  prefix +
  String(
    Math.max(
      0,
      ...arr.map(x =>
        parseInt(String(x[key] || "").replace(/\D/g, "")) || 0
      )
    ) + 1
  ).padStart(5, "0");

const sheet = s => getSheet(s) || [];

// Protect source records once the accounting engine has generated a posted journal.
// Editing/deactivating them without a reversal would leave the ledger inconsistent.
function hasPostedAccountingImpact(record) {
  const sourceMap = {
    Invoices: ["SalesInvoice", "InvoiceID"],
    Bills: ["PurchaseBill", "BillID"],
    CustomerPayments: ["CustomerPayment", "PaymentID"],
    VendorPayments: ["VendorPayment", "PaymentID"],
    InvoiceLines: ["SalesInvoice", "InvoiceID"],
    BillLines: ["PurchaseBill", "BillID"]
  };
  const mapping = sourceMap[cfg.sheet];
  if (!mapping || !record) return false;
  const [sourceType, sourceKey] = mapping;
  const sourceId = record[sourceKey];
  if (sourceId == null || sourceId === "") return false;
  return (getSheet("JournalEntries") || []).some(j =>
    String(j.SourceType) === sourceType && String(j.SourceID) === String(sourceId) &&
    String(j.Status || "Posted").toLowerCase() === "posted"
  );
}

// ==========================================
// MESSAGE
// ==========================================

function message(t, bad = false) {
  $("msg").textContent = t;
  $("msg").className = "msg " + (bad ? "bad" : "good");
}

// ==========================================
// DUPLICATE VALIDATION
// ==========================================

function checkDuplicates(record) {
  const uniqueFields = cfg.uniqueFields || [];

  for (const field of uniqueFields) {
    const value = String(record[field] ?? "")
      .trim()
      .toLowerCase();

    if (!value) continue;

    const duplicate = data.some(row =>
      String(row[field] ?? "")
        .trim()
        .toLowerCase() === value
    );

    if (duplicate) {
      const fieldConfig = cfg.fields.find(f => f.name === field);
      const label = fieldConfig?.label || field;

      throw new Error(
        `${label} already exists. Duplicate records are not allowed.`
      );
    }
  }
}

// ==========================================
// SELECT OPTIONS - FIXED
// ==========================================

function options(field, selectedValue = "") {
  if (!field.options) return "";

  const sourceMap = {
    customers: ["Customers", "CustomerID", "CustomerName"],
    vendors: ["Vendors", "VendorID", "VendorName"],
    items: ["Items", "ItemID", "ItemName"],
    banks: ["Banks", "BankID", "BankName"],
    cash: ["CashAccounts", "CashAccountID", "CashAccountName"],
    invoices: ["Invoices", "InvoiceID", "InvoiceNo"],
    bills: ["Bills", "BillID", "BillNo"],
    coa: ["ChartOfAccounts", "AccountID", "AccountName"]
  };

  const config = sourceMap[field.options];

  if (!config) {
    console.warn("Unknown dropdown source:", field.options);
    return "";
  }

  const [sheetName, defaultValueKey, defaultLabelKey] = config;
  const vals = sheet(sheetName);

  const valueKey = field.valueKey || defaultValueKey;
  const labelKey = field.labelKey || defaultLabelKey;

  return vals
    .filter(x => x && x[valueKey] !== undefined && x[valueKey] !== null)
    .map(x => {
      const value = x[valueKey];
      const label = x[labelKey] || value;
      const selected =
        String(value) === String(selectedValue) ? "selected" : "";

      return `<option value="${esc(value)}" ${selected}>${esc(label)}</option>`;
    })
    .join("");
}

// ==========================================
// INPUT GENERATOR
// ==========================================

function input(f, value = undefined, editMode = false) {
  const v = value !== undefined ? value : (f.value || "");

  if (f.type === "select") {
    const choices = f.options
      ? options(f, v)
      : (f.choices || []).map(x =>
          `<option value="${esc(x)}" ${String(x) === String(v) ? "selected" : ""}>${esc(x)}</option>`
        ).join("");

    return `<select name="${esc(f.name)}" ${f.required ? "required" : ""}>${choices}</select>`;
  }

  return `
    <input
      name="${esc(f.name)}"
      type="${f.type || "text"}"
      value="${esc(v)}"
      ${f.required ? "required" : ""}
      ${f.step ? `step="${f.step}"` : ""}
      ${f.min !== undefined ? `min="${f.min}"` : ""}
    >
  `;
}

// ==========================================
// EDIT MODAL
// ==========================================

function ensureEditModal() {
  if ($("editModal")) return;

  const style = document.createElement("style");

  style.textContent = `
    .acc-modal-backdrop {
      position: fixed;
      inset: 0;
      z-index: 9999;
      display: none;
      align-items: center;
      justify-content: center;
      padding: 20px;
      background: rgba(15, 23, 42, .48);
    }

    .acc-modal {
      width: min(680px, 100%);
      max-height: 90vh;
      overflow: auto;
      background: #fff;
      border-radius: 14px;
      padding: 22px;
      box-shadow: 0 20px 60px rgba(0,0,0,.22);
    }

    .acc-modal h3 {
      margin: 0 0 6px;
      color: #26364e;
      font-size: 19px;
    }

    .acc-modal .acc-modal-desc {
      margin: 0 0 18px;
      color: #8993a5;
      font-size: 13px;
    }

    .acc-modal .acc-modal-actions {
      display: flex;
      justify-content: flex-end;
      gap: 9px;
      margin-top: 20px;
    }

    .acc-modal .secondary {
      background: #eef2f8;
      color: #42516a;
    }

    .acc-modal .modal-error {
      display: none;
      margin-top: 12px;
      padding: 10px;
      border-radius: 7px;
      background: #fff0f0;
      color: #c03939;
      font-size: 12px;
    }

    @media (max-width: 560px) {
      .acc-modal-backdrop {
        padding: 10px;
      }

      .acc-modal {
        padding: 16px;
      }

      .acc-modal .grid {
        grid-template-columns: 1fr;
      }
    }
  `;

  document.head.appendChild(style);

  const backdrop = document.createElement("div");

  backdrop.id = "editModal";
  backdrop.className = "acc-modal-backdrop";

  backdrop.innerHTML = `
    <section class="acc-modal"
      role="dialog"
      aria-modal="true"
      aria-labelledby="editModalTitle">

      <h3 id="editModalTitle">Edit Record</h3>

      <p class="acc-modal-desc">
        Update the record details below.
      </p>

      <form id="editForm">
        <div class="grid" id="editFormFields"></div>

        <div id="editError" class="modal-error"></div>

        <div class="acc-modal-actions">
          <button type="button"
            class="secondary"
            id="cancelEdit">
            Cancel
          </button>

          <button type="submit" id="updateRecordButton">
            Save Changes
          </button>
        </div>
      </form>
    </section>
  `;

  document.body.appendChild(backdrop);

  $("cancelEdit").addEventListener("click", closeEditModal);

  backdrop.addEventListener("click", e => {
    if (e.target === backdrop) closeEditModal();
  });

  $("editForm").addEventListener("submit", saveEditModal);
}

let editingRecord = null;

function closeEditModal() {
  const modal = $("editModal");

  if (modal) modal.style.display = "none";

  editingRecord = null;
}

function openEditModal(record) {
  editingRecord = record;

  ensureEditModal();

  $("editModalTitle").textContent = `Edit ${cfg.title}`;

  $("editFormFields").innerHTML = cfg.fields.map(f => `
    <label>
      ${esc(f.label)}
      ${input(f, record[f.name] ?? "")}
    </label>
  `).join("");

  $("editError").style.display = "none";
  $("editError").textContent = "";

  $("updateRecordButton").disabled = false;
  $("editModal").style.display = "flex";
}

// ==========================================
// SAVE EDIT
// ==========================================

async function saveEditModal(e) {
  e.preventDefault();

  if (!editingRecord) return;

  const button = $("updateRecordButton");
  const errorBox = $("editError");

  try {
    if (!allowed("CanEdit")) {
      throw new Error("No edit permission");
    }

    button.disabled = true;
    errorBox.style.display = "none";

    const fd = new FormData(e.target);
    const changes = {};

    cfg.fields.forEach(f => {
      changes[f.name] = fd.get(f.name);

      if (f.type === "number") {
        changes[f.name] = n(changes[f.name]);
      }
    });

    // Duplicate validation during edit
    for (const field of (cfg.uniqueFields || [])) {
      const value = String(changes[field] ?? "")
        .trim()
        .toLowerCase();

      if (!value) continue;

      const duplicate = data.some(row =>
        String(row[field] ?? "").trim().toLowerCase() === value &&
        String(row[cfg.key]) !== String(editingRecord[cfg.key])
      );

      if (duplicate) {
        const fieldConfig = cfg.fields.find(f => f.name === field);

        throw new Error(
          `${fieldConfig?.label || field} already exists. Duplicate records are not allowed.`
        );
      }
    }

    updateRecord(
      cfg.sheet,
      cfg.key,
      editingRecord[cfg.key],
      withUpdatedAt(changes)
    );

    await saveExcel();

    data = getSheet(cfg.sheet);

    render();
    closeEditModal();

    message("Record updated successfully.");

  } catch (err) {
    console.error("EDIT ERROR:", err);

    errorBox.textContent = err.message;
    errorBox.style.display = "block";

  } finally {
    button.disabled = false;
  }
}

// ==========================================
// RENDER
// ==========================================

function render() {
  $("title").textContent = cfg.title;
  $("desc").textContent = cfg.desc;

  $("formFields").innerHTML = cfg.fields
    .map(f => `
      <label>
        ${esc(f.label)}
        ${input(f)}
      </label>
    `).join("");

  $("save").disabled = !allowed("CanCreate");

  $("tableHead").innerHTML =
    cfg.columns.map(c => `<th>${esc(c.label)}</th>`).join("") +
    "<th>Action</th>";

  $("tableBody").innerHTML = data.map((r, i) => `
    <tr>
      ${cfg.columns.map(c =>
        `<td>${esc(r[c.key])}</td>`
      ).join("")}

      <td>
        ${
          allowed("CanEdit")
            ? `<button data-edit="${i}">Edit</button>`
            : ""
        }

        ${
          allowed("CanDelete")
            ? `<button class="danger" data-del="${i}">Deactivate</button>`
            : ""
        }
      </td>
    </tr>
  `).join("") ||
  `<tr>
    <td colspan="${cfg.columns.length + 1}" class="empty">
      No records found.
    </td>
  </tr>`;
}

// ==========================================
// LOAD DATABASE
// ==========================================

async function load() {
  try {
    console.log("STEP 1: Validating license...");

    const lic = await validateLicense();

    console.log("LICENSE RESULT:", lic);

    if (lic === false || lic?.valid === false) {
      throw new Error(lic?.message || "License invalid");
    }

    console.log("STEP 2: Reading Excel database...");

    const result = await readExcel();

    console.log("WORKBOOK SHEETS:", result);

    console.log("STEP 3: Loading module sheet...");

    data = getSheet(cfg.sheet);

    console.log("MODULE SHEET:", cfg.sheet);
    console.log("RECORD COUNT:", data.length);

    render();

    console.log("STEP 4: Database loaded successfully.");

    message("Database loaded successfully.");

  } catch (e) {
    console.error("DATABASE LOAD ERROR:", e);

    message("Database error: " + e.message, true);
  }
}

// ==========================================
// SAVE NEW RECORD
// ==========================================

$("form").addEventListener("submit", async e => {
  e.preventDefault();

  const saveButton = $("save");

  try {
    if (!allowed("CanCreate")) {
      throw new Error("No create permission");
    }

    saveButton.disabled = true;

    const fd = new FormData(e.target);
    const rec = {};

    cfg.fields.forEach(f => {
      rec[f.name] = fd.get(f.name);
    });

    cfg.fields.forEach(f => {
      if (f.type === "number") {
        rec[f.name] = n(rec[f.name]);
      }
    });

    // Check duplicate records
    checkDuplicates(rec);

    const t = now();

    if (cfg.beforeSave) {
      await cfg.beforeSave(rec);
    }

    rec[cfg.key] = id(cfg.prefix, data, cfg.key);

    const docNo = {
      Invoices: "InvoiceNo",
      Bills: "BillNo",
      CustomerPayments: "PaymentNo",
      VendorPayments: "PaymentNo"
    }[cfg.sheet];

    if (docNo && !rec[docNo]) {
      rec[docNo] = id(cfg.prefix, data, docNo);
    }

    rec.CreatedAt = t;

    if (cfg.updatedAtField !== false) {
      rec.UpdatedAt = t;
    }

    if (cfg.defaults) {
      const defaultValues =
        typeof cfg.defaults === "function"
          ? cfg.defaults(rec)
          : cfg.defaults;

      Object.assign(rec, defaultValues);
    }

    console.log("SAVE: Record prepared:", rec);

    addRecord(cfg.sheet, rec);

    // Post connected accounting effects for source transactions.
    const accountingSheets = ["Invoices", "Bills", "InvoiceLines", "BillLines", "CustomerPayments", "VendorPayments"];
    if (accountingSheets.includes(cfg.sheet)) {
      const { postAccountingTransaction } = await import("./accounting-engine.js");
      await postAccountingTransaction(cfg.sheet, rec);
    }

    if (cfg.afterSave) {
      await cfg.afterSave(rec);
    }

    await saveExcel();

    data = getSheet(cfg.sheet);

    e.target.reset();

    render();

    message("Record saved successfully.");

  } catch (err) {
    console.error("SAVE ERROR:", err);

    message(err.message, true);

  } finally {
    saveButton.disabled = !allowed("CanCreate");
  }
});

// ==========================================
// EDIT AND DEACTIVATE
// ==========================================

$("tableBody").addEventListener("click", async e => {
  const del = e.target.dataset.del;
  const edit = e.target.dataset.edit;

  // DEACTIVATE
  if (del !== undefined) {
    const r = data[Number(del)];

    if (hasPostedAccountingImpact(r)) {
      message("This transaction is already posted to the General Ledger. Create a correcting/reversal transaction instead of deactivating it.", true);
      return;
    }

    if (!confirm("Deactivate this record?")) return;

    try {
      if (!allowed("CanDelete")) {
        throw new Error("No delete permission");
      }

      updateRecord(
        cfg.sheet,
        cfg.key,
        r[cfg.key],
        withUpdatedAt({ IsActive: false })
      );

      await saveExcel();

      data = getSheet(cfg.sheet);

      render();

      message("Record deactivated.");

    } catch (err) {
      console.error("DEACTIVATE ERROR:", err);

      message(err.message, true);
    }
  }

  // EDIT
  else if (edit !== undefined) {
    if (!allowed("CanEdit")) return;

    const r = data[Number(edit)];

    if (hasPostedAccountingImpact(r)) {
      message("This transaction is already posted to the General Ledger. Editing is blocked to protect the journal. Use a correcting/reversal transaction instead.", true);
      return;
    }

    openEditModal(r);
  }
});

// ==========================================
// INITIALIZE
// ==========================================

load();
