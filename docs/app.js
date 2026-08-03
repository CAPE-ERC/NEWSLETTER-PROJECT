// Static public viewer — no backend. Everything here is either bundled JSON/
// Markdown served alongside this page, or kept in the browser (localStorage).
// See src/cape_newsletter/agents/writer_agents.py:write_section for the real
// (API-key-backed) equivalent this page reconstructs prompts for.

const DATA_DIR = "./data";
const PROMPTS_DIR = "./prompts";

// Mirrors orchestrator.SECTIONS in src/cape_newsletter/agents/orchestrator.py.
const SECTIONS = [
  "highlights",
  "global_economic_update",
  "global_economic_outlook",
  "nigeria_output_growth",
  "output_growth_outlook",
  "price_update",
  "fiscal_operations",
  "conclusion",
  "country_in_focus",
];

// Mirrors SECTION_DATA_KEY in src/cape_newsletter/api.py. Sections absent
// here (highlights, conclusion) have no single structured-data block — the
// builder falls back to the full edition record for those.
const SECTION_DATA_KEY = {
  global_economic_update: "global",
  global_economic_outlook: "global",
  nigeria_output_growth: "nigeria_output",
  output_growth_outlook: "nigeria_output",
  price_update: "prices",
  fiscal_operations: "fiscal",
  country_in_focus: "country_in_focus",
};

// Mirrors READY_SECTIONS in src/cape_newsletter/api.py.
const READY_SECTIONS = new Set(["price_update"]);

const MANUAL_FIELD_HINTS = {
  country: "Editorial choice — supplied manually each month, not inferred from data.",
};

const editionRecordCache = new Map();
const promptTextCache = new Map();
let savedBuilderData = null;
const bDataFieldInputs = new Map();

async function fetchJSON(url) {
  const res = await fetch(url);
  if (!res.ok) throw new Error(`${url}: ${res.status}`);
  return res.json();
}

async function fetchText(url) {
  if (!promptTextCache.has(url)) {
    const res = await fetch(url);
    if (!res.ok) throw new Error(`${url}: ${res.status}`);
    promptTextCache.set(url, await res.text());
  }
  return promptTextCache.get(url);
}

async function getEditionRecord(id) {
  if (!editionRecordCache.has(id)) {
    editionRecordCache.set(id, await fetchJSON(`${DATA_DIR}/processed/${id}.json`));
  }
  return editionRecordCache.get(id);
}

// --- editions explorer -------------------------------------------------------

function renderEditionBlock(title, obj) {
  if (!obj) return "";
  const rows = Object.entries(obj)
    .filter(([k]) => k !== "notes")
    .map(([k, v]) => `<dt>${k}</dt><dd>${v === null ? "—" : v}</dd>`)
    .join("");
  const notes = obj.notes ? `<p class="notes">${obj.notes}</p>` : "";
  return `<h4>${title}</h4><dl>${rows}</dl>${notes}`;
}

async function selectEdition(id, listEl) {
  [...listEl.children].forEach((li) => li.classList.toggle("active", li.dataset.id === id));
  const record = await getEditionRecord(id);
  const detail = document.getElementById("edition-detail");

  detail.innerHTML = [
    renderEditionBlock("Nigeria output", record.nigeria_output),
    renderEditionBlock("Prices", record.prices),
    renderEditionBlock("Fiscal", record.fiscal),
    renderEditionBlock("Country in focus", record.country_in_focus),
  ].join("");

  if (record.extraction_notes) {
    const p = document.createElement("p");
    p.className = "notes";
    p.textContent = record.extraction_notes;
    detail.appendChild(p);
  }
}

async function loadEditions() {
  const editions = await fetchJSON(`${DATA_DIR}/editions.json`);
  const list = document.getElementById("edition-list");
  const bSelect = document.getElementById("b-edition");
  list.innerHTML = "";
  bSelect.innerHTML = "";

  editions.forEach((ed, i) => {
    const li = document.createElement("li");
    li.textContent = `${ed.month} ${ed.year}`;
    li.dataset.id = ed.id;
    if (i === editions.length - 1) li.classList.add("active");
    li.addEventListener("click", () => selectEdition(ed.id, list));
    list.appendChild(li);

    const opt = document.createElement("option");
    opt.value = ed.id;
    opt.textContent = `${ed.month} ${ed.year}`;
    bSelect.appendChild(opt);
  });

  if (editions.length) await selectEdition(editions[editions.length - 1].id, list);
}

function loadSections() {
  const select = document.getElementById("b-section");
  select.innerHTML = "";
  SECTIONS.forEach((name) => {
    const opt = document.createElement("option");
    opt.value = name;
    opt.textContent = READY_SECTIONS.has(name)
      ? name.replace(/_/g, " ")
      : `${name.replace(/_/g, " ")} (placeholder)`;
    select.appendChild(opt);
  });
}

// --- data/questions panel (mirrors web/app.js) --------------------------------

function renderDataFields(data) {
  const container = document.getElementById("b-data-fields");
  container.innerHTML = "";
  bDataFieldInputs.clear();

  Object.entries(data).forEach(([key, value]) => {
    const wrap = document.createElement("label");
    wrap.className = "data-field";

    const labelText = document.createElement("span");
    labelText.className = "data-field-label";
    labelText.textContent = key.replace(/_/g, " ");
    wrap.appendChild(labelText);

    if (MANUAL_FIELD_HINTS[key]) {
      const hint = document.createElement("span");
      hint.className = "data-field-hint";
      hint.textContent = MANUAL_FIELD_HINTS[key];
      wrap.appendChild(hint);
    }

    const isObject = value !== null && typeof value === "object";
    const asText = isObject ? JSON.stringify(value) : value ?? "";
    const long = typeof asText === "string" && asText.length > 60;

    const input = document.createElement(long ? "textarea" : "input");
    if (!long) input.type = "text";
    input.value = asText;
    input.dataset.key = key;
    input.dataset.type = isObject ? "object" : typeof value;
    wrap.appendChild(input);

    bDataFieldInputs.set(key, input);
    container.appendChild(wrap);
  });
}

function collectDataFromFields() {
  const result = {};
  bDataFieldInputs.forEach((input, key) => {
    const raw = input.value.trim();
    const type = input.dataset.type;
    if (raw === "") {
      result[key] = null;
    } else if (type === "number") {
      const n = Number(raw);
      result[key] = Number.isNaN(n) ? raw : n;
    } else if (type === "object") {
      try {
        result[key] = JSON.parse(raw);
      } catch {
        result[key] = raw;
      }
    } else {
      result[key] = raw;
    }
  });
  return result;
}

async function updateDataPanel() {
  const edition_id = document.getElementById("b-edition").value;
  const section = document.getElementById("b-section").value;
  const readinessEl = document.getElementById("b-readiness");
  document.getElementById("b-prompt-out").hidden = true;

  readinessEl.textContent = READY_SECTIONS.has(section) ? "ready" : "placeholder template";
  readinessEl.className = `readiness-badge ${READY_SECTIONS.has(section) ? "ready" : "placeholder"}`;

  if (!edition_id) return;
  const record = await getEditionRecord(edition_id);
  const dataKey = SECTION_DATA_KEY[section];

  if (dataKey) {
    savedBuilderData = record[dataKey] || {};
  } else {
    // highlights / conclusion: no dedicated block, fall back to the full record.
    const { extraction_notes, ...rest } = record;
    savedBuilderData = rest;
  }
  renderDataFields(savedBuilderData);
  loadOutputForCurrentSelection();
}

// --- prompt builder ------------------------------------------------------------

async function buildPrompt() {
  const section = document.getElementById("b-section").value;
  const sectionMd = await fetchText(`${PROMPTS_DIR}/sections/${section}.md`);
  const styleGuideMd = await fetchText(`${PROMPTS_DIR}/style_guide.md`);
  const system = `${sectionMd}\n\n---\n\n${styleGuideMd}`;

  const data = collectDataFromFields();
  const user =
    "Draft this section from the following structured data. Do not state any figure that is not present in this JSON:\n\n" +
    JSON.stringify(data, null, 2);

  document.getElementById("b-system").value = system;
  document.getElementById("b-user").value = user;
  document.getElementById("b-prompt-out").hidden = false;
}

async function copyText(text) {
  try {
    await navigator.clipboard.writeText(text);
    return true;
  } catch {
    return false;
  }
}

// --- output paste-back (localStorage only, never sent anywhere) --------------

function outputStorageKey() {
  const edition_id = document.getElementById("b-edition").value;
  const section = document.getElementById("b-section").value;
  return `cape-dashboard-output:${edition_id}:${section}`;
}

function loadOutputForCurrentSelection() {
  const saved = localStorage.getItem(outputStorageKey());
  document.getElementById("b-output").value = saved || "";
}

// --- wiring --------------------------------------------------------------------

document.getElementById("b-edition").addEventListener("change", updateDataPanel);
document.getElementById("b-section").addEventListener("change", updateDataPanel);
document.getElementById("b-data-reset").addEventListener("click", () => {
  if (savedBuilderData) renderDataFields(savedBuilderData);
});
document.getElementById("b-build").addEventListener("click", () => {
  buildPrompt().catch((err) => {
    document.getElementById("b-prompt-out").hidden = false;
    document.getElementById("b-system").value = `Could not load prompt template: ${err.message}`;
    document.getElementById("b-user").value = "";
  });
});

document.querySelectorAll(".copy-button").forEach((btn) => {
  btn.addEventListener("click", async () => {
    const target = document.getElementById(btn.dataset.copyTarget);
    const ok = await copyText(target.value);
    btn.textContent = ok ? "Copied" : "Copy failed";
    setTimeout(() => (btn.textContent = "Copy"), 1200);
  });
});

document.getElementById("b-copy-both").addEventListener("click", async (evt) => {
  const system = document.getElementById("b-system").value;
  const user = document.getElementById("b-user").value;
  const combined = `SYSTEM:\n${system}\n\nUSER:\n${user}`;
  const ok = await copyText(combined);
  evt.target.textContent = ok ? "Copied" : "Copy failed";
  setTimeout(() => (evt.target.textContent = "Copy both (ready to paste into claude.ai)"), 1200);
});

document.getElementById("b-output").addEventListener("input", (evt) => {
  localStorage.setItem(outputStorageKey(), evt.target.value);
});
document.getElementById("b-clear-output").addEventListener("click", () => {
  localStorage.removeItem(outputStorageKey());
  document.getElementById("b-output").value = "";
});

async function init() {
  await loadEditions();
  loadSections();
  await updateDataPanel();
}

init();
