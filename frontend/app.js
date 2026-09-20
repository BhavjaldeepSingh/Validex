// ==========================================================================
// TRANSACTION VALIDATION PLATFORM - IN-BROWSER LOCAL STORAGE ENGINE
// ==========================================================================

// const API_URL =
    "https://transaction-validator-api.onrender.com";

const API_URL = "https://validex-xp3x.onrender.com";

const LOCAL_STORAGE_KEY = "tv_validation_history_v2";

// Built-in phone validation rules: {country_code: {prefix: string, length: number}}
const PHONE_RULES = {
    "IN": { prefix: "91", length: 12, localLength: 10 },
    "SG": { prefix: "65", length: 10, localLength: 8 },
    "US": { prefix: "1", length: 11, localLength: 10 },
    "CA": { prefix: "1", length: 11, localLength: 10 },
    "UK": { prefix: "44", length: 12, localLength: 10 },
    "GB": { prefix: "44", length: 12, localLength: 10 },
    "AU": { prefix: "61", length: 11, localLength: 9 },
    "AE": { prefix: "971", length: 12, localLength: 9 },
    "DE": { prefix: "49", length: 12, localLength: 10 },
    "FR": { prefix: "33", length: 11, localLength: 9 },
};

const DEFAULT_PAYMENT_MODES = [
    "CARD", "COD", "PAYPAL", "NETBANKING", "WALLET", "UPI", "BANK_TRANSFER", "CRYPTO", "STRIPE"
];

const DEFAULT_PAYMENT_STATUSES = [
    "COMPLETED", "SUCCESS", "PENDING", "FAILED", "REFUNDED", "CANCELLED", "PROCESSING"
];

// App State
let currentDataset = [];
let currentColumns = [];
let currentFilename = "dataset.csv";
let currentWorkbook = null;
let currentValidationResult = null;

// Pagination State
let errorPage = 1;
let cleanPage = 1;
const PAGE_SIZE = 15;
let filteredErrors = [];
let filteredClean = [];

// ==========================================================================
// INITIALIZATION
// ==========================================================================

window.addEventListener("DOMContentLoaded", () => {
    initDropzone();
    initToolbarButtons();
    initTabs();
    initPaginationControls();
    initExportButtons();
    loadLocalStorageHistory();
    updateStorageMeter();
});

// ==========================================================================
// FILE INGESTION & DROPZONE
// ==========================================================================

function initDropzone() {
    const dropzone = document.getElementById("dropzone");
    const fileInput = document.getElementById("fileInput");
    const browseFileBtn = document.getElementById("browseFileBtn");
    const reselectFileBtn = document.getElementById("reselectFileBtn");
    const loadSampleBtn = document.getElementById("loadSampleBtn");

    browseFileBtn.addEventListener("click", () => fileInput.click());
    reselectFileBtn.addEventListener("click", () => fileInput.click());

    fileInput.addEventListener("change", (e) => {
        if (e.target.files && e.target.files[0]) {
            processSelectedFile(e.target.files[0]);
        }
    });

    dropzone.addEventListener("dragover", (e) => {
        e.preventDefault();
        dropzone.classList.add("drag-over");
    });

    dropzone.addEventListener("dragleave", () => {
        dropzone.classList.remove("drag-over");
    });

    dropzone.addEventListener("drop", (e) => {
        e.preventDefault();
        dropzone.classList.remove("drag-over");
        if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
            processSelectedFile(e.dataTransfer.files[0]);
        }
    });

    loadSampleBtn.addEventListener("click", loadSampleDataset);
}

function processSelectedFile(file) {
    currentFilename = file.name;
    const ext = file.name.split(".").pop().toLowerCase();

    const reader = new FileReader();

    if (ext === "xlsx" || ext === "xls") {
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            if (typeof XLSX !== "undefined") {
                currentWorkbook = XLSX.read(data, { type: "array" });
                setupExcelSheetSelector(currentWorkbook);
            } else {
                showToast("Excel parser library loading, please wait...", "error");
            }
        };
        reader.readAsArrayBuffer(file);
    } else if (ext === "json") {
        reader.onload = (e) => {
            try {
                const parsed = JSON.parse(e.target.result);
                const records = Array.isArray(parsed) ? parsed : [parsed];
                onDatasetLoaded(records, file.name);
            } catch (err) {
                showToast("Invalid JSON file: " + err.message, "error");
            }
        };
        reader.readAsText(file);
    } else {
        // CSV or TSV
        reader.onload = (e) => {
            const text = e.target.result;
            const delimiter = ext === "tsv" ? "\t" : undefined;
            if (typeof Papa !== "undefined") {
                Papa.parse(text, {
                    header: true,
                    skipEmptyLines: true,
                    dynamicTyping: false,
                    delimiter: delimiter,
                    complete: (results) => {
                        onDatasetLoaded(results.data, file.name);
                    },
                    error: (err) => {
                        showToast("Error parsing CSV: " + err.message, "error");
                    }
                });
            } else {
                // Pure JS fallback CSV parser
                const records = parseCsvFallback(text, delimiter || ",");
                onDatasetLoaded(records, file.name);
            }
        };
        reader.readAsText(file);
    }
}

function setupExcelSheetSelector(workbook) {
    const sheetSelect = document.getElementById("excelSheetSelect");
    const sheetWrap = document.getElementById("sheetSelectorWrap");
    sheetSelect.innerHTML = "";

    workbook.SheetNames.forEach((name, idx) => {
        const opt = document.createElement("option");
        opt.value = name;
        opt.textContent = name;
        if (idx === 0) opt.selected = true;
        sheetSelect.appendChild(opt);
    });

    sheetWrap.style.display = "flex";

    sheetSelect.onchange = () => {
        loadExcelSheet(sheetSelect.value);
    };

    loadExcelSheet(workbook.SheetNames[0]);
}

function loadExcelSheet(sheetName) {
    if (!currentWorkbook) return;
    const worksheet = currentWorkbook.Sheets[sheetName];
    const records = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
    onDatasetLoaded(records, `${currentFilename} [${sheetName}]`);
}

// Pure JS CSV Parser fallback (RFC 4180 compliant)
function parseCsvFallback(text, delimiter = ",") {
    const lines = text.split(/\r\n|\n|\r/);
    if (!lines.length) return [];
    
    // Parse header
    const headers = parseCsvLine(lines[0], delimiter);
    const records = [];

    for (let i = 1; i < lines.length; i++) {
        const line = lines[i].trim();
        if (!line) continue;
        const values = parseCsvLine(line, delimiter);
        const row = {};
        headers.forEach((h, idx) => {
            row[h] = values[idx] !== undefined ? values[idx] : "";
        });
        records.push(row);
    }
    return records;
}

function parseCsvLine(line, delimiter) {
    const result = [];
    let cur = "";
    let inQuotes = false;
    for (let i = 0; i < line.length; i++) {
        const ch = line[i];
        if (ch === '"') {
            if (inQuotes && line[i + 1] === '"') {
                cur += '"';
                i++;
            } else {
                inQuotes = !inQuotes;
            }
        } else if (ch === delimiter && !inQuotes) {
            result.push(cur.trim());
            cur = "";
        } else {
            cur += ch;
        }
    }
    result.push(cur.trim());
    return result;
}

// ==========================================================================
// DATASET LOADED & SCHEMA DISCOVERY
// ==========================================================================

function onDatasetLoaded(records, filename) {
    if (!records || records.length === 0) {
        showToast("The selected dataset is empty.", "error");
        return;
    }

    currentDataset = records;
    currentFilename = filename;

    // Discover column names
    const colSet = new Set();
    records.slice(0, 50).forEach(row => {
        Object.keys(row).forEach(k => {
            if (k && !k.startsWith("__")) colSet.add(k);
        });
    });
    currentColumns = Array.from(colSet);

    // Update active file banner
    document.getElementById("activeFileName").textContent = filename;
    document.getElementById("activeFileMeta").textContent = `${records.length.toLocaleString()} rows detected • ${currentColumns.length} columns`;
    document.getElementById("activeFileBanner").style.display = "flex";

    // Show Constraints Panel
    document.getElementById("constraintsPanel").style.display = "block";
    renderCustomConstraintsStudio(currentColumns, records);

    // Smooth scroll to constraints
    document.getElementById("constraintsPanel").scrollIntoView({ behavior: "smooth", block: "start" });
    showToast(`Loaded ${records.length} records. Review or adjust column constraints.`, "success");
}

// ==========================================================================
// SMART COLUMN RULES INFERENCE
// ==========================================================================

function inferColumnRule(colName, sampleRows) {
    const colLower = colName.toLowerCase().trim();
    const tokens = colLower.split(/[\s_\-]+/);
    const samples = sampleRows.map(r => String(r[colName] || "").trim()).filter(Boolean);

    let inferredType = "string";
    let inferredRole = "general";
    let isRequired = true;
    let isUnique = false;
    let minVal = null;
    let maxVal = null;
    let allowedValues = "";
    let dateFormat = "DD-MM-YYYY";
    let phoneCountry = "IN";
    let regexPattern = "";

    // 1. Check Unique ID using token boundary to avoid "amount_paid" matching "id"
    const isPrimaryId = 
        colLower === "id" || colLower === "uuid" ||
        colLower === "transaction_id" || colLower === "txn_id" ||
        colLower === "order_id" || colLower === "invoice_id" || colLower === "invoice_no" ||
        ((tokens.includes("transaction") || tokens.includes("order") || tokens.includes("invoice") || tokens.includes("txn")) && tokens.includes("id")) ||
        (colLower.endsWith("_id") && !["product", "item", "category", "status", "country", "customer", "user", "paid"].some(k => tokens.includes(k)));

    if (isPrimaryId) {
        inferredRole = "unique_id";
        isUnique = true;
        isRequired = true;
        inferredType = "string";
    }
    // 2. Check Phone Number
    else if (tokens.some(t => ["phone", "mobile", "contact", "cell", "tel", "whatsapp"].includes(t)) || colLower.includes("phone") || colLower.includes("mobile")) {
        inferredRole = "phone_number";
        isRequired = true;
        inferredType = "string";
    }
    // 3. Check Country Code
    else if (tokens.some(t => ["country", "nation", "geo", "region"].includes(t))) {
        inferredRole = "country_code";
        isRequired = true;
        inferredType = "string";
    }
    // 4. Check Payment Mode
    else if (colLower.includes("payment_mode") || colLower.includes("pay_mode") || tokens.includes("mode") || tokens.includes("gateway")) {
        inferredRole = "payment_mode";
        allowedValues = DEFAULT_PAYMENT_MODES.join(", ");
        isRequired = true;
        inferredType = "string";
    }
    // 5. Check Payment Status
    else if (colLower.includes("payment_status") || colLower.includes("txn_status") || (tokens.includes("status") && !tokens.includes("product"))) {
        inferredRole = "payment_status";
        allowedValues = DEFAULT_PAYMENT_STATUSES.join(", ");
        isRequired = true;
        inferredType = "string";
    }
    // 6. Check Date
    else if (tokens.some(t => ["date", "dob", "created_at", "dt"].includes(t)) || colLower.includes("date")) {
        inferredRole = "date";
        inferredType = "date";
        isRequired = true;
        dateFormat = "DD-MM-YYYY";
    }
    // 7. Check Time
    else if (tokens.some(t => ["time", "hour", "slot"].includes(t))) {
        inferredRole = "time";
        inferredType = "time";
        isRequired = false;
    }
    // 8. Check Amount / Price
    else if (tokens.some(t => ["amount", "price", "total", "cost", "fee", "amt", "paid", "balance", "revenue"].includes(t))) {
        inferredRole = "amount";
        inferredType = "float";
        minVal = 0;
        isRequired = true;
    }
    // 9. Check Quantity
    else if (tokens.some(t => ["quantity", "qty", "units", "count", "items"].includes(t))) {
        inferredRole = "quantity";
        inferredType = "integer";
        minVal = 1;
        isRequired = true;
    }
    // 10. Check Email
    else if (colLower.includes("email") || colLower.includes("mail")) {
        inferredRole = "email";
        isRequired = true;
    }
    // 11. Generic Check by Value inspection
    else {
        let allNumbers = samples.length > 0 && samples.every(s => !isNaN(Number(s.replace(/[^0-9.-]+/g, ""))));
        let allIntegers = allNumbers && samples.every(s => Number.isInteger(Number(s.replace(/[^0-9.-]+/g, ""))));
        let datePattern = samples.length > 0 && samples.every(s => /\d{1,4}[-/\.]\d{1,2}[-/\.]\d{1,4}/.test(s));

        if (datePattern) {
            inferredType = "date";
            inferredRole = "date";
        } else if (allIntegers) {
            inferredType = "integer";
            minVal = 0;
        } else if (allNumbers) {
            inferredType = "float";
            minVal = 0;
        } else {
            inferredType = "string";
            inferredRole = "general";
        }
    }

    return {
        column_name: colName,
        data_type: inferredType,
        semantic_role: inferredRole,
        is_required: isRequired,
        is_unique: isUnique,
        min_value: minVal,
        max_value: maxVal,
        allowed_values: allowedValues,
        date_format: dateFormat,
        default_country: phoneCountry,
        regex_pattern: regexPattern,
        enabled: true
    };
}

// ==========================================================================
// RENDER CUSTOM CONSTRAINTS STUDIO
// ==========================================================================

function renderCustomConstraintsStudio(columns, records) {
    const grid = document.getElementById("columnGrid");
    grid.innerHTML = "";

    const sampleRows = records.slice(0, 15);

    columns.forEach(col => {
        const rule = inferColumnRule(col, sampleRows);
        const samples = sampleRows.map(r => r[col]).filter(v => v !== undefined && v !== "").slice(0, 3);
        const card = createColumnConstraintCard(col, rule, samples);
        grid.appendChild(card);
    });

    updateSelectedColumnsCounter();
}

function createColumnConstraintCard(colName, rule, samples) {
    const card = document.createElement("div");
    card.className = "column-card";
    card.dataset.column = colName;

    const safeColId = colName.replace(/[^a-zA-Z0-9_]/g, "_");

    card.innerHTML = `
        <div>
            <!-- Header -->
            <div class="column-card-header">
                <div class="column-header-left">
                    <input type="checkbox" class="col-toggle-checkbox" id="chk_enabled_${safeColId}" checked title="Validate this column">
                    <span class="column-name-title" title="${colName}">${colName}</span>
                </div>
                <span class="col-role-badge" id="badge_role_${safeColId}">${rule.semantic_role}</span>
            </div>

            <!-- Samples -->
            <div class="col-sample-preview">
                <span style="font-size:0.72rem; font-weight:600; color:var(--text-muted);">Samples:</span>
                ${samples.length ? samples.map(s => `<span class="sample-pill">${escapeHtml(String(s))}</span>`).join("") : '<span class="sample-pill">&lt;empty&gt;</span>'}
            </div>

            <!-- Constraint Controls -->
            <div class="constraint-controls" id="controls_${safeColId}">
                <div class="control-row">
                    <div class="input-label-pair">
                        <label>Data Type</label>
                        <select class="glass-select col-data-type" id="type_${safeColId}">
                            <option value="string" ${rule.data_type === "string" ? "selected" : ""}>String / Text</option>
                            <option value="integer" ${rule.data_type === "integer" ? "selected" : ""}>Integer</option>
                            <option value="float" ${rule.data_type === "float" ? "selected" : ""}>Float / Decimal</option>
                            <option value="date" ${rule.data_type === "date" ? "selected" : ""}>Date</option>
                            <option value="time" ${rule.data_type === "time" ? "selected" : ""}>Time</option>
                            <option value="boolean" ${rule.data_type === "boolean" ? "selected" : ""}>Boolean</option>
                        </select>
                    </div>

                    <div class="input-label-pair">
                        <label>Semantic Role</label>
                        <select class="glass-select col-role" id="role_${safeColId}">
                            <option value="general" ${rule.semantic_role === "general" ? "selected" : ""}>General</option>
                            <option value="unique_id" ${rule.semantic_role === "unique_id" ? "selected" : ""}>Unique ID</option>
                            <option value="phone_number" ${rule.semantic_role === "phone_number" ? "selected" : ""}>Phone Number</option>
                            <option value="amount" ${rule.semantic_role === "amount" ? "selected" : ""}>Amount / Price</option>
                            <option value="quantity" ${rule.semantic_role === "quantity" ? "selected" : ""}>Quantity</option>
                            <option value="payment_mode" ${rule.semantic_role === "payment_mode" ? "selected" : ""}>Payment Mode</option>
                            <option value="payment_status" ${rule.semantic_role === "payment_status" ? "selected" : ""}>Payment Status</option>
                            <option value="date" ${rule.semantic_role === "date" ? "selected" : ""}>Date</option>
                            <option value="time" ${rule.semantic_role === "time" ? "selected" : ""}>Time</option>
                            <option value="email" ${rule.semantic_role === "email" ? "selected" : ""}>Email</option>
                            <option value="country_code" ${rule.semantic_role === "country_code" ? "selected" : ""}>Country Code</option>
                        </select>
                    </div>
                </div>

                <div class="control-row">
                    <div class="toggle-group">
                        <label for="req_${safeColId}">Required</label>
                        <label class="switch">
                            <input type="checkbox" class="col-is-required" id="req_${safeColId}" ${rule.is_required ? "checked" : ""}>
                            <span class="slider"></span>
                        </label>
                    </div>

                    <div class="toggle-group">
                        <label for="uniq_${safeColId}">Unique</label>
                        <label class="switch">
                            <input type="checkbox" class="col-is-unique" id="uniq_${safeColId}" ${rule.is_unique ? "checked" : ""}>
                            <span class="slider"></span>
                        </label>
                    </div>
                </div>

                <!-- Bounds (Numeric) -->
                <div class="control-row numeric-bounds-row" id="bounds_${safeColId}" style="display: ${['integer', 'float'].includes(rule.data_type) || ['amount', 'quantity'].includes(rule.semantic_role) ? 'grid' : 'none'};">
                    <div class="input-label-pair">
                        <label>Min Value</label>
                        <input type="number" step="any" class="glass-input col-min-val" id="min_${safeColId}" placeholder="None" value="${rule.min_value !== null ? rule.min_value : ""}">
                    </div>
                    <div class="input-label-pair">
                        <label>Max Value</label>
                        <input type="number" step="any" class="glass-input col-max-val" id="max_${safeColId}" placeholder="None" value="${rule.max_value !== null ? rule.max_value : ""}">
                    </div>
                </div>

                <!-- Phone Settings -->
                <div class="control-row phone-settings-row" id="phone_${safeColId}" style="display: ${rule.semantic_role === 'phone_number' ? 'grid' : 'none'};">
                    <div class="input-label-pair" style="grid-column: 1 / -1;">
                        <label>Default Country Rule</label>
                        <select class="glass-select col-phone-country" id="country_${safeColId}">
                            <option value="IN" ${rule.default_country === "IN" ? "selected" : ""}>India (IN - 10 / 12 digits)</option>
                            <option value="US" ${rule.default_country === "US" ? "selected" : ""}>United States / CA (+1 - 10/11 digits)</option>
                            <option value="UK" ${rule.default_country === "UK" ? "selected" : ""}>United Kingdom (+44 - 10/12 digits)</option>
                            <option value="SG" ${rule.default_country === "SG" ? "selected" : ""}>Singapore (+65 - 8/10 digits)</option>
                            <option value="AE" ${rule.default_country === "AE" ? "selected" : ""}>UAE (+971 - 9/12 digits)</option>
                            <option value="AU" ${rule.default_country === "AU" ? "selected" : ""}>Australia (+61 - 9/11 digits)</option>
                        </select>
                    </div>
                </div>

                <!-- Date Settings -->
                <div class="input-label-pair date-settings-row" id="date_${safeColId}" style="display: ${rule.data_type === 'date' || rule.semantic_role === 'date' ? 'flex' : 'none'};">
                    <label>Accepted Date Format</label>
                    <select class="glass-select col-date-format" id="dfmt_${safeColId}">
                        <option value="DD-MM-YYYY">DD-MM-YYYY (e.g. 29-05-2026)</option>
                        <option value="YYYY-MM-DD">YYYY-MM-DD (e.g. 2026-05-29)</option>
                        <option value="MM/DD/YYYY">MM/DD/YYYY (e.g. 05/29/2026)</option>
                        <option value="DD/MM/YYYY">DD/MM/YYYY (e.g. 29/05/2026)</option>
                        <option value="ANY">Any Standard Date Format</option>
                    </select>
                </div>

                <!-- Allowed Values (Enum) -->
                <div class="input-label-pair enum-settings-row" id="enum_${safeColId}" style="display: ${['payment_mode', 'payment_status'].includes(rule.semantic_role) ? 'flex' : 'none'};">
                    <label>Allowed Values (Comma-separated)</label>
                    <input type="text" class="glass-input col-allowed-vals" id="allowed_${safeColId}" placeholder="CARD, UPI, COD" value="${rule.allowed_values}">
                </div>

                <!-- Custom Regex -->
                <div class="input-label-pair">
                    <label>Custom Regex (Optional)</label>
                    <input type="text" class="glass-input col-regex" id="regex_${safeColId}" placeholder="e.g. ^ORD\\d+$" value="${rule.regex_pattern}">
                </div>
            </div>
        </div>
    `;

    // Dynamic UI change listeners
    const toggleEnabled = card.querySelector(`#chk_enabled_${safeColId}`);
    const roleSelect = card.querySelector(`#role_${safeColId}`);
    const typeSelect = card.querySelector(`#type_${safeColId}`);
    const roleBadge = card.querySelector(`#badge_role_${safeColId}`);
    const boundsRow = card.querySelector(`#bounds_${safeColId}`);
    const phoneRow = card.querySelector(`#phone_${safeColId}`);
    const dateRow = card.querySelector(`#date_${safeColId}`);
    const enumRow = card.querySelector(`#enum_${safeColId}`);
    const isUniqSwitch = card.querySelector(`#uniq_${safeColId}`);

    toggleEnabled.addEventListener("change", (e) => {
        card.classList.toggle("disabled", !e.target.checked);
        updateSelectedColumnsCounter();
    });

    roleSelect.addEventListener("change", (e) => {
        const val = e.target.value;
        roleBadge.textContent = val;

        // Auto sync type & inputs
        if (val === "unique_id") {
            isUniqSwitch.checked = true;
        } else if (val === "amount") {
            typeSelect.value = "float";
            boundsRow.style.display = "grid";
        } else if (val === "quantity") {
            typeSelect.value = "integer";
            boundsRow.style.display = "grid";
        } else if (val === "date") {
            typeSelect.value = "date";
        } else if (val === "payment_mode" && !card.querySelector(`#allowed_${safeColId}`).value) {
            card.querySelector(`#allowed_${safeColId}`).value = DEFAULT_PAYMENT_MODES.join(", ");
        } else if (val === "payment_status" && !card.querySelector(`#allowed_${safeColId}`).value) {
            card.querySelector(`#allowed_${safeColId}`).value = DEFAULT_PAYMENT_STATUSES.join(", ");
        }

        phoneRow.style.display = val === "phone_number" ? "grid" : "none";
        enumRow.style.display = ["payment_mode", "payment_status"].includes(val) ? "flex" : "none";
        dateRow.style.display = (val === "date" || typeSelect.value === "date") ? "flex" : "none";
    });

    typeSelect.addEventListener("change", (e) => {
        const val = e.target.value;
        boundsRow.style.display = ["integer", "float"].includes(val) ? "grid" : "none";
        dateRow.style.display = (val === "date" || roleSelect.value === "date") ? "flex" : "none";
    });

    return card;
}

function updateSelectedColumnsCounter() {
    const total = document.querySelectorAll(".column-card").length;
    const enabled = document.querySelectorAll(".col-toggle-checkbox:checked").length;
    document.getElementById("columnsSelectedCount").textContent = `${enabled} of ${total} columns active`;
}

// ==========================================================================
// PRESETS & TOOLBAR
// ==========================================================================

function initToolbarButtons() {
    document.getElementById("presetSmartBtn").addEventListener("click", () => {
        renderCustomConstraintsStudio(currentColumns, currentDataset);
        showToast("Auto-inferred smart rules applied.", "success");
    });

    document.getElementById("presetStrictBtn").addEventListener("click", applyStrictFinancialPreset);

    document.getElementById("toggleAllColsBtn").addEventListener("click", (e) => {
        const checkboxes = document.querySelectorAll(".col-toggle-checkbox");
        const allChecked = Array.from(checkboxes).every(c => c.checked);
        checkboxes.forEach(c => {
            c.checked = !allChecked;
            c.closest(".column-card").classList.toggle("disabled", allChecked);
        });
        e.target.textContent = allChecked ? "Select All" : "Deselect All";
        updateSelectedColumnsCounter();
    });

    document.getElementById("resetRulesBtn").addEventListener("click", () => {
        renderCustomConstraintsStudio(currentColumns, currentDataset);
        showToast("Constraints reset to default.", "success");
    });

    document.getElementById("runValidationBtn").addEventListener("click", executeInBrowserValidation);
    document.getElementById("revalidateBtn").addEventListener("click", executeInBrowserValidation);
    document.getElementById("clearHistoryBtn").addEventListener("click", clearLocalStorageHistory);
}

function applyStrictFinancialPreset() {
    document.querySelectorAll(".column-card").forEach(card => {
        const col = card.dataset.column;
        const colLower = col.toLowerCase();
        const safeColId = col.replace(/[^a-zA-Z0-9_]/g, "_");

        const reqChk = card.querySelector(`#req_${safeColId}`);
        const uniqChk = card.querySelector(`#uniq_${safeColId}`);
        const minVal = card.querySelector(`#min_${safeColId}`);
        const role = card.querySelector(`#role_${safeColId}`).value;

        if (reqChk) reqChk.checked = true;

        if (role === "unique_id" || colLower.includes("order_id") || colLower.includes("txn_id") || colLower.includes("transaction_id")) {
            if (uniqChk) uniqChk.checked = true;
        }

        if (role === "amount" || colLower.includes("amount") || colLower.includes("price")) {
            if (minVal) minVal.value = "0";
        }
    });
    showToast("Strict Financial Audit Preset Applied.", "success");
}

// ==========================================================================
// IN-BROWSER VALIDATION ENGINE
// ==========================================================================

function collectRulesFromUI() {
    const rules = [];
    const countryColRef = currentColumns.find(c => {
        const cl = c.toLowerCase();
        return cl.includes("country") || cl === "geo";
    });

    document.querySelectorAll(".column-card").forEach(card => {
        const colName = card.dataset.column;
        const safeColId = colName.replace(/[^a-zA-Z0-9_]/g, "_");

        const isEnabled = card.querySelector(`#chk_enabled_${safeColId}`).checked;
        if (!isEnabled) return;

        const dataType = card.querySelector(`#type_${safeColId}`).value;
        const semanticRole = card.querySelector(`#role_${safeColId}`).value;
        const isRequired = card.querySelector(`#req_${safeColId}`).checked;
        const isUnique = card.querySelector(`#uniq_${safeColId}`).checked;
        const minValRaw = card.querySelector(`#min_${safeColId}`).value;
        const maxValRaw = card.querySelector(`#max_${safeColId}`).value;
        const allowedValsRaw = card.querySelector(`#allowed_${safeColId}`).value;
        const dateFormat = card.querySelector(`#dfmt_${safeColId}`).value;
        const defaultCountry = card.querySelector(`#country_${safeColId}`) ? card.querySelector(`#country_${safeColId}`).value : "IN";
        const regexPattern = card.querySelector(`#regex_${safeColId}`).value.trim();

        const allowedList = allowedValsRaw ? allowedValsRaw.split(",").map(s => s.trim().toUpperCase()).filter(Boolean) : [];

        rules.push({
            column_name: colName,
            data_type: dataType,
            semantic_role: semanticRole,
            is_required: isRequired,
            is_unique: isUnique,
            min_value: minValRaw !== "" ? parseFloat(minValRaw) : null,
            max_value: maxValRaw !== "" ? parseFloat(maxValRaw) : null,
            allowed_values: allowedList,
            date_format: dateFormat,
            default_country: defaultCountry,
            country_column_ref: countryColRef,
            regex_pattern: regexPattern
        });
    });

    return rules;
}

function executeInBrowserValidation() {
    if (!currentDataset || currentDataset.length === 0) {
        showToast("Please upload or load a dataset first.", "error");
        return;
    }

    const rules = collectRulesFromUI();
    if (rules.length === 0) {
        showToast("Please enable at least one column constraint.", "error");
        return;
    }

    const runBtn = document.getElementById("runValidationBtn");
    runBtn.disabled = true;
    runBtn.textContent = "Validating in Browser...";

    setTimeout(() => {
        try {
            const result = performValidation(currentDataset, rules);
            currentValidationResult = result;

            // Save run to local storage
            saveRunToLocalStorage(result);

            // Render Results Dashboard
            renderResultsDashboard(result);

            // Update History Table & Storage Meter
            loadLocalStorageHistory();
            updateStorageMeter();

            showToast("In-browser validation completed successfully!", "success");
        } catch (err) {
            console.error("Validation error:", err);
            showToast("Validation failed: " + err.message, "error");
        } finally {
            runBtn.disabled = false;
            runBtn.innerHTML = `
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
                    <polygon points="5 3 19 12 5 21 5 3"></polygon>
                </svg>
                Run In-Browser Validation
            `;
        }
    }, 50);
}

function performValidation(dataset, rules) {
    const cleanRecords = [];
    const errorRecords = [];
    const errorBreakdown = {
        "Missing Required Field": 0,
        "Duplicate Unique Key": 0,
        "Type / Parsing Error": 0,
        "Numeric Boundary Violation": 0,
        "Invalid Phone Number": 0,
        "Invalid Date / Time": 0,
        "Invalid Enum Value": 0,
        "Format / Regex Failure": 0,
    };

    // Tracking seen sets per unique column
    const seenUniqueSets = {};
    rules.filter(r => r.is_unique).forEach(r => {
        seenUniqueSets[r.column_name] = new Set();
    });

    const enableChunking = document.getElementById("enableChunking").checked;
    const chunkSize = parseInt(document.getElementById("chunkSizeInput").value, 10) || 100;

    dataset.forEach((row, index) => {
        const rowNumber = index + 2; // 1-based row with header
        const rowErrors = [];
        const cellErrors = {};

        for (const rule of rules) {
            const col = rule.column_name;
            const val = row[col];
            const valStr = (val !== null && val !== undefined) ? String(val).trim() : "";
            const isEmpty = valStr === "" || valStr.toLowerCase() === "nan" || valStr.toLowerCase() === "null";

            // 1. Required Check
            if (isEmpty) {
                if (rule.is_required) {
                    const msg = `Required column '${col}' is empty.`;
                    rowErrors.push(msg);
                    cellErrors[col] = { reason: msg, type: "Missing Required Field", val: "<EMPTY>" };
                    errorBreakdown["Missing Required Field"]++;
                }
                continue;
            }

            // 2. Uniqueness Check
            if (rule.is_unique) {
                if (seenUniqueSets[col].has(valStr)) {
                    const msg = `Duplicate key '${valStr}' in unique column '${col}'.`;
                    rowErrors.push(msg);
                    cellErrors[col] = { reason: msg, type: "Duplicate Unique Key", val: valStr };
                    errorBreakdown["Duplicate Unique Key"]++;
                } else {
                    seenUniqueSets[col].add(valStr);
                }
            }

            // 3. Numeric & Bounds Check
            if (["integer", "float"].includes(rule.data_type) || ["amount", "quantity"].includes(rule.semantic_role)) {
                const cleanedNumStr = valStr.replace(/[^0-9.-]/g, "");
                const numVal = parseFloat(cleanedNumStr);

                if (isNaN(numVal) || cleanedNumStr === "") {
                    const msg = `Value '${valStr}' is not a valid ${rule.data_type}.`;
                    rowErrors.push(msg);
                    cellErrors[col] = { reason: msg, type: "Type / Parsing Error", val: valStr };
                    errorBreakdown["Type / Parsing Error"]++;
                } else {
                    if (rule.min_value !== null && numVal < rule.min_value) {
                        const msg = `Value ${numVal} is below minimum ${rule.min_value}.`;
                        rowErrors.push(msg);
                        cellErrors[col] = { reason: msg, type: "Numeric Boundary Violation", val: valStr };
                        errorBreakdown["Numeric Boundary Violation"]++;
                    }
                    if (rule.max_value !== null && numVal > rule.max_value) {
                        const msg = `Value ${numVal} exceeds maximum ${rule.max_value}.`;
                        rowErrors.push(msg);
                        cellErrors[col] = { reason: msg, type: "Numeric Boundary Violation", val: valStr };
                        errorBreakdown["Numeric Boundary Violation"]++;
                    }
                }
            }

            // 4. Date & Time Validation
            if (rule.data_type === "date" || rule.semantic_role === "date") {
                if (!isValidDateString(valStr, rule.date_format)) {
                    const msg = `Date '${valStr}' does not match format '${rule.date_format}'.`;
                    rowErrors.push(msg);
                    cellErrors[col] = { reason: msg, type: "Invalid Date / Time", val: valStr };
                    errorBreakdown["Invalid Date / Time"]++;
                }
            }

            // 5. Phone Number Validation
            if (rule.semantic_role === "phone_number") {
                let country = rule.default_country || "IN";
                if (rule.country_column_ref && row[rule.country_column_ref]) {
                    const rowCountry = String(row[rule.country_column_ref]).trim().toUpperCase();
                    if (PHONE_RULES[rowCountry]) country = rowCountry;
                }

                const digits = valStr.replace(/\D/g, "");
                const phoneRule = PHONE_RULES[country];

                let phoneValid = false;
                if (phoneRule) {
                    // Match with international prefix or local length
                    const startsWithPrefix = digits.startsWith(phoneRule.prefix) && digits.length === phoneRule.length;
                    const matchesLocal = digits.length === phoneRule.localLength;
                    phoneValid = startsWithPrefix || matchesLocal;
                } else {
                    phoneValid = digits.length >= 8 && digits.length <= 15;
                }

                if (!phoneValid) {
                    const msg = `Phone '${valStr}' is invalid for country '${country}'.`;
                    rowErrors.push(msg);
                    cellErrors[col] = { reason: msg, type: "Invalid Phone Number", val: valStr };
                    errorBreakdown["Invalid Phone Number"]++;
                }
            }

            // 6. Email Validation
            if (rule.semantic_role === "email") {
                const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
                if (!emailRegex.test(valStr)) {
                    const msg = `Invalid email address '${valStr}'.`;
                    rowErrors.push(msg);
                    cellErrors[col] = { reason: msg, type: "Format / Regex Failure", val: valStr };
                    errorBreakdown["Format / Regex Failure"]++;
                }
            }

            // 7. Allowed Values (Enum) Check
            if (rule.allowed_values && rule.allowed_values.length > 0) {
                if (!rule.allowed_values.includes(valStr.toUpperCase())) {
                    const msg = `Value '${valStr}' not permitted. Expected: ${rule.allowed_values.slice(0, 4).join(", ")}.`;
                    rowErrors.push(msg);
                    cellErrors[col] = { reason: msg, type: "Invalid Enum Value", val: valStr };
                    errorBreakdown["Invalid Enum Value"]++;
                }
            }

            // 8. Custom Regex Check
            if (rule.regex_pattern) {
                try {
                    const reg = new RegExp(rule.regex_pattern);
                    if (!reg.test(valStr)) {
                        const msg = `Value '${valStr}' failed custom regex '${rule.regex_pattern}'.`;
                        rowErrors.push(msg);
                        cellErrors[col] = { reason: msg, type: "Format / Regex Failure", val: valStr };
                        errorBreakdown["Format / Regex Failure"]++;
                    }
                } catch (e) {
                    // Ignore regex syntax errors
                }
            }
        }

        // Record classification
        if (rowErrors.length === 0) {
            cleanRecords.push(row);
        } else {
            const errRecord = Object.assign({}, row);
            errRecord["__error_row_number__"] = rowNumber;
            errRecord["__validation_errors__"] = rowErrors.join(" | ");
            errRecord["__cell_errors__"] = cellErrors;
            errorRecords.push(errRecord);
        }
    });

    // Partition into Chunks
    const chunks = [];
    if (enableChunking && chunkSize > 0) {
        for (let i = 0; i < cleanRecords.length; i += chunkSize) {
            chunks.push({
                chunk_index: Math.floor(i / chunkSize) + 1,
                name: `chunk_${Math.floor(i / chunkSize) + 1}.csv`,
                start_row: i + 1,
                end_row: Math.min(i + chunkSize, cleanRecords.length),
                records: cleanRecords.slice(i, i + chunkSize)
            });
        }
    }

    const totalRows = dataset.length;
    const validRows = cleanRecords.length;
    const invalidRows = errorRecords.length;
    const healthScore = totalRows > 0 ? Math.round((validRows / totalRows) * 100) : 100;

    return {
        id: "run_" + Date.now(),
        filename: currentFilename,
        timestamp: new Date().toISOString(),
        total_rows: totalRows,
        valid_rows: validRows,
        invalid_rows: invalidRows,
        health_score: healthScore,
        error_breakdown: errorBreakdown,
        clean_records: cleanRecords,
        error_records: errorRecords,
        chunks: chunks,
        columns: currentColumns,
        rule_count: rules.length
    };
}

// Robust Date String Validator
function isValidDateString(dateStr, format) {
    if (!dateStr) return false;
    
    // Test general date
    if (format === "ANY") {
        const d = new Date(dateStr);
        return !isNaN(d.getTime());
    }

    if (format === "DD-MM-YYYY" || format === "DD/MM/YYYY") {
        const parts = dateStr.split(/[-/.]/);
        if (parts.length !== 3) return false;
        const day = parseInt(parts[0], 10);
        const month = parseInt(parts[1], 10);
        const year = parseInt(parts[2], 10);
        return day >= 1 && day <= 31 && month >= 1 && month <= 12 && year >= 1900 && year <= 2100;
    }

    if (format === "YYYY-MM-DD") {
        const parts = dateStr.split(/[-/.]/);
        if (parts.length !== 3) return false;
        const year = parseInt(parts[0], 10);
        const month = parseInt(parts[1], 10);
        const day = parseInt(parts[2], 10);
        return day >= 1 && day <= 31 && month >= 1 && month <= 12 && year >= 1900 && year <= 2100;
    }

    if (format === "MM/DD/YYYY") {
        const parts = dateStr.split(/[-/.]/);
        if (parts.length !== 3) return false;
        const month = parseInt(parts[0], 10);
        const day = parseInt(parts[1], 10);
        const year = parseInt(parts[2], 10);
        return day >= 1 && day <= 31 && month >= 1 && month <= 12 && year >= 1900 && year <= 2100;
    }

    const d = new Date(dateStr);
    return !isNaN(d.getTime());
}

// ==========================================================================
// RENDER RESULTS DASHBOARD
// ==========================================================================

function renderResultsDashboard(result) {
    const resultsPanel = document.getElementById("resultsPanel");
    resultsPanel.style.display = "block";

    // KPIs
    document.getElementById("kpiTotal").textContent = result.total_rows.toLocaleString();
    document.getElementById("kpiValid").textContent = result.valid_rows.toLocaleString();
    document.getElementById("kpiInvalid").textContent = result.invalid_rows.toLocaleString();
    document.getElementById("kpiScore").textContent = `${result.health_score}%`;

    const validPct = result.total_rows > 0 ? Math.round((result.valid_rows / result.total_rows) * 100) : 0;
    const invalidPct = result.total_rows > 0 ? Math.round((result.invalid_rows / result.total_rows) * 100) : 0;

    document.getElementById("kpiValidPercent").textContent = `${validPct}% pass rate`;
    document.getElementById("kpiInvalidPercent").textContent = `${invalidPct}% rejected`;

    // Breakdown Chips
    const breakdownContainer = document.getElementById("breakdownChips");
    breakdownContainer.innerHTML = "";
    let hasErrors = false;

    for (const [category, count] of Object.entries(result.error_breakdown)) {
        if (count > 0) {
            hasErrors = true;
            const chip = document.createElement("span");
            chip.className = "error-chip";
            chip.innerHTML = `${category} <span class="error-chip-count">${count}</span>`;
            breakdownContainer.appendChild(chip);
        }
    }

    if (!hasErrors) {
        breakdownContainer.innerHTML = `<span style="font-size:0.85rem; color:#059669; font-weight:600;">✨ Zero errors found! All records passed validation cleanly.</span>`;
    }

    // Update Tab Badges
    document.getElementById("tabErrorsBadge").textContent = result.invalid_rows;
    document.getElementById("tabCleanBadge").textContent = result.valid_rows;
    document.getElementById("tabChunksBadge").textContent = result.chunks.length;

    // Filter states
    filteredErrors = result.error_records;
    filteredClean = result.clean_records;
    errorPage = 1;
    cleanPage = 1;

    renderErrorsTable();
    renderCleanTable();
    renderChunksGrid(result.chunks);

    // If there are errors, open Errors tab, else open Clean tab
    const defaultTab = result.invalid_rows > 0 ? "tabErrors" : "tabClean";
    activateTab(defaultTab);

    resultsPanel.scrollIntoView({ behavior: "smooth", block: "start" });
}

// ==========================================================================
// RENDER DATA TABLES (ERRORS & CLEAN)
// ==========================================================================

function renderErrorsTable() {
    const thead = document.getElementById("errorTableHead");
    const tbody = document.getElementById("errorTableBody");
    const pageInfo = document.getElementById("errorPageInfo");

    thead.innerHTML = "";
    tbody.innerHTML = "";

    if (!filteredErrors || filteredErrors.length === 0) {
        tbody.innerHTML = `<tr><td colspan="${currentColumns.length + 2}" class="empty-state">No error records found matching criteria.</td></tr>`;
        pageInfo.textContent = "Page 1 of 1";
        return;
    }

    // Build Header
    const trHead = document.createElement("tr");
    trHead.innerHTML = `<th>Row #</th><th>Error Diagnostics</th>` + currentColumns.map(c => `<th>${c}</th>`).join("");
    thead.appendChild(trHead);

    // Slice for pagination
    const totalPages = Math.ceil(filteredErrors.length / PAGE_SIZE) || 1;
    if (errorPage > totalPages) errorPage = totalPages;
    const start = (errorPage - 1) * PAGE_SIZE;
    const pageRecords = filteredErrors.slice(start, start + PAGE_SIZE);

    pageRecords.forEach(row => {
        const tr = document.createElement("tr");
        const cellErrors = row["__cell_errors__"] || {};

        let rowHtml = `
            <td style="font-weight:700; color:var(--text-secondary);">${row["__error_row_number__"]}</td>
            <td style="color:#b91c1c; font-size:0.8rem; font-weight:600; max-width:320px; white-space:normal;">
                ${escapeHtml(row["__validation_errors__"] || "")}
            </td>
        `;

        currentColumns.forEach(col => {
            const val = row[col] !== undefined ? String(row[col]) : "";
            const errInfo = cellErrors[col];

            if (errInfo) {
                rowHtml += `
                    <td class="cell-error" title="${escapeHtml(errInfo.reason)}">
                        ${escapeHtml(val || "<EMPTY>")}
                        <span class="error-badge-pill">!</span>
                    </td>
                `;
            } else {
                rowHtml += `<td>${escapeHtml(val)}</td>`;
            }
        });

        tr.innerHTML = rowHtml;
        tbody.appendChild(tr);
    });

    pageInfo.textContent = `Page ${errorPage} of ${totalPages} (${filteredErrors.length} records)`;
    document.getElementById("prevErrorPage").disabled = errorPage <= 1;
    document.getElementById("nextErrorPage").disabled = errorPage >= totalPages;
}

function renderCleanTable() {
    const thead = document.getElementById("cleanTableHead");
    const tbody = document.getElementById("cleanTableBody");
    const pageInfo = document.getElementById("cleanPageInfo");

    thead.innerHTML = "";
    tbody.innerHTML = "";

    if (!filteredClean || filteredClean.length === 0) {
        tbody.innerHTML = `<tr><td colspan="${currentColumns.length}" class="empty-state">No clean records found.</td></tr>`;
        pageInfo.textContent = "Page 1 of 1";
        return;
    }

    // Build Header
    const trHead = document.createElement("tr");
    trHead.innerHTML = currentColumns.map(c => `<th>${c}</th>`).join("");
    thead.appendChild(trHead);

    // Slice for pagination
    const totalPages = Math.ceil(filteredClean.length / PAGE_SIZE) || 1;
    if (cleanPage > totalPages) cleanPage = totalPages;
    const start = (cleanPage - 1) * PAGE_SIZE;
    const pageRecords = filteredClean.slice(start, start + PAGE_SIZE);

    pageRecords.forEach(row => {
        const tr = document.createElement("tr");
        tr.innerHTML = currentColumns.map(c => `<td>${escapeHtml(row[c] !== undefined ? String(row[c]) : "")}</td>`).join("");
        tbody.appendChild(tr);
    });

    pageInfo.textContent = `Page ${cleanPage} of ${totalPages} (${filteredClean.length} records)`;
    document.getElementById("prevCleanPage").disabled = cleanPage <= 1;
    document.getElementById("nextCleanPage").disabled = cleanPage >= totalPages;
}

function renderChunksGrid(chunks) {
    const grid = document.getElementById("chunksGrid");
    grid.innerHTML = "";

    if (!chunks || chunks.length === 0) {
        grid.innerHTML = `<div class="empty-state" style="grid-column:1/-1;">Chunking disabled or zero valid records to partition.</div>`;
        return;
    }

    chunks.forEach(chunk => {
        const card = document.createElement("div");
        card.className = "chunk-card";
        card.innerHTML = `
            <div class="chunk-card-title">
                <span>${chunk.name}</span>
                <span class="badge-version">${chunk.records.length} rows</span>
            </div>
            <div class="chunk-card-meta">
                Rows ${chunk.start_row} to ${chunk.end_row}
            </div>
            <button type="button" class="btn btn-secondary btn-sm" style="width:100%;">
                Inspect Chunk in Browser
            </button>
        `;

        card.querySelector("button").addEventListener("click", () => {
            previewChunkTable(chunk);
        });

        grid.appendChild(card);
    });
}

function previewChunkTable(chunk) {
    const wrap = document.getElementById("chunkRecordsWrap");
    const title = document.getElementById("chunkRecordsTitle");
    const thead = document.getElementById("chunkTableHead");
    const tbody = document.getElementById("chunkTableBody");

    title.textContent = `Previewing ${chunk.name} (${chunk.records.length} records, Rows ${chunk.start_row} - ${chunk.end_row})`;
    thead.innerHTML = "<tr>" + currentColumns.map(c => `<th>${c}</th>`).join("") + "</tr>";
    tbody.innerHTML = "";

    chunk.records.slice(0, 50).forEach(r => {
        const tr = document.createElement("tr");
        tr.innerHTML = currentColumns.map(c => `<td>${escapeHtml(r[c] !== undefined ? String(r[c]) : "")}</td>`).join("");
        tbody.appendChild(tr);
    });

    wrap.style.display = "block";
    wrap.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

// ==========================================================================
// TABS & SEARCH
// ==========================================================================

function initTabs() {
    document.querySelectorAll(".tab-btn").forEach(btn => {
        btn.addEventListener("click", () => {
            activateTab(btn.dataset.tab);
        });
    });

    // Errors Search
    document.getElementById("searchErrorsInput").addEventListener("input", (e) => {
        const query = e.target.value.toLowerCase().trim();
        if (!currentValidationResult) return;
        if (!query) {
            filteredErrors = currentValidationResult.error_records;
        } else {
            filteredErrors = currentValidationResult.error_records.filter(row => {
                return Object.values(row).some(v => String(v).toLowerCase().includes(query));
            });
        }
        errorPage = 1;
        renderErrorsTable();
    });

    // Clean Search
    document.getElementById("searchCleanInput").addEventListener("input", (e) => {
        const query = e.target.value.toLowerCase().trim();
        if (!currentValidationResult) return;
        if (!query) {
            filteredClean = currentValidationResult.clean_records;
        } else {
            filteredClean = currentValidationResult.clean_records.filter(row => {
                return Object.values(row).some(v => String(v).toLowerCase().includes(query));
            });
        }
        cleanPage = 1;
        renderCleanTable();
    });
}

function activateTab(tabId) {
    document.querySelectorAll(".tab-btn").forEach(b => {
        b.classList.toggle("active", b.dataset.tab === tabId);
    });
    document.querySelectorAll(".tab-pane").forEach(p => {
        p.style.display = p.id === tabId ? "block" : "none";
    });
}

function initPaginationControls() {
    document.getElementById("prevErrorPage").addEventListener("click", () => {
        if (errorPage > 1) {
            errorPage--;
            renderErrorsTable();
        }
    });

    document.getElementById("nextErrorPage").addEventListener("click", () => {
        const totalPages = Math.ceil(filteredErrors.length / PAGE_SIZE);
        if (errorPage < totalPages) {
            errorPage++;
            renderErrorsTable();
        }
    });

    document.getElementById("prevCleanPage").addEventListener("click", () => {
        if (cleanPage > 1) {
            cleanPage--;
            renderCleanTable();
        }
    });

    document.getElementById("nextCleanPage").addEventListener("click", () => {
        const totalPages = Math.ceil(filteredClean.length / PAGE_SIZE);
        if (cleanPage < totalPages) {
            cleanPage++;
            renderCleanTable();
        }
    });
}

// ==========================================================================
// LOCAL STORAGE PERSISTENCE & HISTORY
// ==========================================================================

function saveRunToLocalStorage(result) {
    try {
        const existingRaw = localStorage.getItem(LOCAL_STORAGE_KEY);
        const history = existingRaw ? JSON.parse(existingRaw) : [];

        // Store lightweight summary with records (capped if massive)
        const runItem = {
            id: result.id,
            filename: result.filename,
            timestamp: result.timestamp,
            total_rows: result.total_rows,
            valid_rows: result.valid_rows,
            invalid_rows: result.invalid_rows,
            health_score: result.health_score,
            error_breakdown: result.error_breakdown,
            columns: result.columns,
            chunk_count: result.chunks.length,
            // Keep records in storage
            clean_records: result.clean_records.slice(0, 500),
            error_records: result.error_records.slice(0, 500),
            chunks: result.chunks.map(c => ({
                chunk_index: c.chunk_index,
                name: c.name,
                start_row: c.start_row,
                end_row: c.end_row,
                records: c.records.slice(0, 200)
            }))
        };

        // Prepend and cap history to latest 15 runs to maintain healthy localStorage
        history.unshift(runItem);
        if (history.length > 15) history.pop();

        localStorage.setItem(LOCAL_STORAGE_KEY, JSON.stringify(history));
    } catch (e) {
        console.warn("LocalStorage save warning:", e);
        showToast("Storage quota near limit. Older history pruned.", "error");
    }
}

function loadLocalStorageHistory() {
    const tbody = document.getElementById("historyBody");
    tbody.innerHTML = "";

    try {
        const existingRaw = localStorage.getItem(LOCAL_STORAGE_KEY);
        const history = existingRaw ? JSON.parse(existingRaw) : [];

        if (!history || history.length === 0) {
            tbody.innerHTML = `<tr><td colspan="8" class="empty-state">No validation history stored yet. Run a validation to persist sessions locally.</td></tr>`;
            return;
        }

        history.forEach(item => {
            const tr = document.createElement("tr");
            tr.innerHTML = `
                <td style="font-family:'JetBrains Mono', monospace; font-size:0.8rem; font-weight:700;">${item.id}</td>
                <td style="font-weight:600;">${escapeHtml(item.filename)}</td>
                <td>${item.total_rows.toLocaleString()}</td>
                <td style="color:#059669; font-weight:700;">${item.valid_rows.toLocaleString()}</td>
                <td style="color:#e11d48; font-weight:700;">${item.invalid_rows.toLocaleString()}</td>
                <td><span class="badge-version">${item.health_score || 100}%</span></td>
                <td style="font-size:0.8rem; color:var(--text-secondary);">${new Date(item.timestamp).toLocaleString()}</td>
                <td>
                    <button type="button" class="btn btn-secondary btn-sm inspect-hist-btn" data-id="${item.id}">Inspect</button>
                    <button type="button" class="btn btn-secondary btn-sm delete-hist-btn" data-id="${item.id}" style="color:#e11d48;">✕</button>
                </td>
            `;

            tr.querySelector(".inspect-hist-btn").addEventListener("click", () => {
                restoreRunFromHistory(item);
            });

            tr.querySelector(".delete-hist-btn").addEventListener("click", () => {
                deleteRunFromHistory(item.id);
            });

            tbody.appendChild(tr);
        });
    } catch (e) {
        console.error("Error loading history:", e);
    }
}

function restoreRunFromHistory(runItem) {
    currentValidationResult = runItem;
    currentColumns = runItem.columns || [];
    currentFilename = runItem.filename;

    renderResultsDashboard(runItem);
    showToast(`Loaded session ${runItem.id} from local storage.`, "success");
}

function deleteRunFromHistory(id) {
    try {
        const raw = localStorage.getItem(LOCAL_STORAGE_KEY);
        if (!raw) return;
        let history = JSON.parse(raw);
        history = history.filter(item => item.id !== id);
        localStorage.setItem(LOCAL_STORAGE_KEY, JSON.stringify(history));
        loadLocalStorageHistory();
        updateStorageMeter();
        showToast(`Run deleted from local storage.`, "success");
    } catch (e) {
        console.error(e);
    }
}

function clearLocalStorageHistory() {
    if (confirm("Are you sure you want to clear all validation history stored in your browser?")) {
        localStorage.removeItem(LOCAL_STORAGE_KEY);
        loadLocalStorageHistory();
        updateStorageMeter();
        showToast("Local storage history cleared.", "success");
    }
}

function updateStorageMeter() {
    try {
        let totalBytes = 0;
        for (let key in localStorage) {
            if (localStorage.hasOwnProperty(key)) {
                totalBytes += (localStorage[key].length + key.length) * 2;
            }
        }
        const kb = Math.round(totalBytes / 1024);
        document.getElementById("storageUsageText").textContent = `${kb} KB / 5 MB`;
        const pct = Math.min(100, Math.round((kb / 5120) * 100));
        document.getElementById("storageMeterFill").style.width = `${pct}%`;
    } catch (e) {
        // Ignore
    }
}

// ==========================================================================
// OPTIONAL EXPORT (CLIENT-SIDE)
// ==========================================================================

function initExportButtons() {
    document.getElementById("exportCleanCsvBtn").addEventListener("click", () => {
        if (!currentValidationResult || !currentValidationResult.clean_records.length) {
            showToast("No clean records available to export.", "error");
            return;
        }
        downloadCsvFile(currentValidationResult.clean_records, `${currentFilename.replace(/\.[^/.]+$/, "")}_clean.csv`);
    });

    document.getElementById("exportErrorCsvBtn").addEventListener("click", () => {
        if (!currentValidationResult || !currentValidationResult.error_records.length) {
            showToast("No error records to export.", "error");
            return;
        }
        downloadCsvFile(currentValidationResult.error_records, `${currentFilename.replace(/\.[^/.]+$/, "")}_errors.csv`);
    });

    document.getElementById("exportAllJsonBtn").addEventListener("click", () => {
        if (!currentValidationResult) {
            showToast("No validation result available to export.", "error");
            return;
        }
        const jsonStr = JSON.stringify(currentValidationResult, null, 2);
        downloadBlob(new Blob([jsonStr], { type: "application/json" }), `${currentFilename.replace(/\.[^/.]+$/, "")}_audit.json`);
    });
}

function downloadCsvFile(records, filename) {
    if (!records || !records.length) return;
    const keys = Object.keys(records[0]).filter(k => k !== "__cell_errors__");
    let csv = keys.map(k => `"${k}"`).join(",") + "\n";

    records.forEach(row => {
        const line = keys.map(k => {
            const v = row[k] !== undefined ? String(row[k]) : "";
            return `"${v.replace(/"/g, '""')}"`;
        }).join(",");
        csv += line + "\n";
    });

    downloadBlob(new Blob([csv], { type: "text/csv;charset=utf-8;" }), filename);
    showToast(`Exported ${filename}`, "success");
}

function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

// ==========================================================================
// SAMPLE DATASET GENERATOR
// ==========================================================================

function loadSampleDataset() {
    const sampleData = [
        { order_id: "ORD20260001", order_date: "29-05-2026", customer_name: "Diya Das", phone_number: "+65 95674573", country: "SG", product_id: "P001", product_name: "Wireless Mechanical Keyboard", quantity: "5", unit_price: "89.99", payment_mode: "CARD", payment_status: "FAILED", amount_paid: "0.0", transaction_id: "TXN20260001" },
        { order_id: "ORD20260002", order_date: "26-02-2026", customer_name: "Rohan Joshi", phone_number: "+1 7912151271", country: "US", product_id: "P003", product_name: "4K Ultra HD Monitor 27", quantity: "4", unit_price: "329.99", payment_mode: "CARD", payment_status: "COMPLETED", amount_paid: "1319.96", transaction_id: "TXN20260002" },
        { order_id: "ORD20260003", order_date: "19-02-2026", customer_name: "Ananya Verma", phone_number: "+65 98029735", country: "SG", product_id: "P003", product_name: "4K Ultra HD Monitor 27", quantity: "2", unit_price: "329.99", payment_mode: "BITCOIN", payment_status: "COMPLETED", amount_paid: "659.98", transaction_id: "TXN20260003" }, // Invalid mode
        { order_id: "ORD20260004", order_date: "03-03-2026", customer_name: "Michael Gupta", phone_number: "+91 7270148938", country: "IN", product_id: "P004", product_name: "Noise Cancelling Headphones", quantity: "4", unit_price: "199.99", payment_mode: "PAYPAL", payment_status: "COMPLETED", amount_paid: "799.96", transaction_id: "TXN20260004" },
        { order_id: "ORD20260005", order_date: "35-01-2026", customer_name: "Neha Mehta", phone_number: "+1 8148124468", country: "US", product_id: "P004", product_name: "Noise Cancelling Headphones", quantity: "4", unit_price: "199.99", payment_mode: "CARD", payment_status: "PENDING", amount_paid: "-50.0", transaction_id: "TXN20260005" }, // Invalid date & negative amount
        { order_id: "ORD20260006", order_date: "20-01-2026", customer_name: "Michael Gupta", phone_number: "+91 8993317548", country: "IN", product_id: "P002", product_name: "Ergonomic Wireless Mouse", quantity: "5", unit_price: "49.99", payment_mode: "CARD", payment_status: "COMPLETED", amount_paid: "249.95", transaction_id: "TXN20260006" },
        { order_id: "ORD20260007", order_date: "06-01-2026", customer_name: "Emily Mehta", phone_number: "+91 123", country: "IN", product_id: "P005", product_name: "USB-C Dual Docking Station", quantity: "2", unit_price: "129.99", payment_mode: "CARD", payment_status: "PENDING", amount_paid: "0.0", transaction_id: "TXN20260007" }, // Invalid phone
        { order_id: "ORD20260008", order_date: "16-02-2026", customer_name: "Meera Verma", phone_number: "+91 8738064505", country: "IN", product_id: "P003", product_name: "4K Ultra HD Monitor 27", quantity: "5", unit_price: "329.99", payment_mode: "COD", payment_status: "COMPLETED", amount_paid: "1649.95", transaction_id: "TXN20260008" },
        { order_id: "ORD20260009", order_date: "08-05-2026", customer_name: "Rahul Patel", phone_number: "+91 9968167226", country: "IN", product_id: "P005", product_name: "USB-C Dual Docking Station", quantity: "1", unit_price: "129.99", payment_mode: "CARD", payment_status: "COMPLETED", amount_paid: "129.99", transaction_id: "TXN20260009" },
        { order_id: "ORD20260002", order_date: "09-01-2026", customer_name: "Aman Mehta", phone_number: "+1 8899525898", country: "US", product_id: "P001", product_name: "Wireless Mechanical Keyboard", quantity: "5", unit_price: "89.99", payment_mode: "NETBANKING", payment_status: "PENDING", amount_paid: "0.0", transaction_id: "TXN20260010" }, // Duplicate order_id (ORD20260002)
        { order_id: "ORD20260011", order_date: "05-02-2026", customer_name: "Ananya Kumar", phone_number: "+1 9350788890", country: "US", product_id: "P004", product_name: "Noise Cancelling Headphones", quantity: "2", unit_price: "199.99", payment_mode: "CARD", payment_status: "COMPLETED", amount_paid: "399.98", transaction_id: "TXN20260011" },
        { order_id: "ORD20260012", order_date: "21-03-2026", customer_name: "", phone_number: "+91 7255887634", country: "IN", product_id: "P005", product_name: "USB-C Dual Docking Station", quantity: "5", unit_price: "129.99", payment_mode: "WALLET", payment_status: "COMPLETED", amount_paid: "649.95", transaction_id: "TXN20260012" } // Empty required customer_name
    ];

    onDatasetLoaded(sampleData, "sample_transactions_audit.csv");
}

// ==========================================================================
// UTILITIES
// ==========================================================================

function escapeHtml(str) {
    if (str === null || str === undefined) return "";
    return String(str)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

function showToast(message, type = "info") {
    const container = document.getElementById("toastContainer");
    const toast = document.createElement("div");
    toast.className = `toast ${type}`;

    const icon = type === "success" ? "✓" : (type === "error" ? "⚠" : "ℹ");
    toast.innerHTML = `<span style="font-size:1.1rem;">${icon}</span><span>${escapeHtml(message)}</span>`;

    container.appendChild(toast);

    setTimeout(() => {
        toast.style.opacity = "0";
        toast.style.transform = "translateX(100%)";
        toast.style.transition = "all 0.3s ease";
        setTimeout(() => toast.remove(), 300);
    }, 4000);
}
