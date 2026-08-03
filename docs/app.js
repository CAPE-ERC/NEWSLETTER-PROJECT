// CAPE Newsletter — Agent Workbench (static, no backend).
//
// Mirrors the pipeline in docs/architecture.md / src/cape_newsletter/agents/:
//   5 collect agents -> 9 draft (writer) agents -> 1 chart agent -> 1 edit agent.
// Each agent step builds the exact prompt that agent would receive, from a
// file parsed client-side, another agent's saved output, or pasted text. The
// model call itself happens outside this page (claude.ai, or the real
// Anthropic-backed pipeline) — the output is pasted back in and persisted to
// localStorage so downstream agents can pull it as input.

const PROMPTS_DIR = "./prompts";
const SCHEMAS_DIR = "./schemas";
const DATA_DIR = "./data";
const STORAGE_KEY = "cape-workbench-v1";

const PHASES = [
  { key: "collect", label: "Collect", color: "var(--phase-collect)" },
  { key: "draft", label: "Draft", color: "var(--phase-draft)" },
  { key: "chart", label: "Chart", color: "var(--phase-chart)" },
  { key: "edit", label: "Edit", color: "var(--phase-edit)" },
];

// Mirrors data_agents.py's five fetch_* functions.
const AGENTS = {
  collect: [
    { key: "global_data_agent", label: "Global Data Agent", schemaKey: "global", sourceHint: "IMF / World Bank / market commentary (report, PDF, or pasted text)" },
    { key: "nigeria_real_sector_agent", label: "Nigeria Real-Sector Agent", schemaKey: "nigeria_output", sourceHint: "Stanbic IBTC PMI report, NBS GDP release" },
    { key: "price_cpi_agent", label: "Price / CPI Agent", schemaKey: "prices", sourceHint: "NBS CPI report" },
    { key: "fiscal_faac_agent", label: "Fiscal / FAAC Agent", schemaKey: "fiscal", sourceHint: "FAAC disbursement circular" },
    { key: "country_focus_agent", label: "Country-Focus Agent", schemaKey: "country_in_focus", sourceHint: "CAPE's own EPU index/research note" },
  ],
  // Mirrors orchestrator.SECTIONS order. `upstream` names the collect
  // agent(s) this section's data would naturally come from.
  draft: [
    { key: "highlights", label: "Highlights", upstream: ["collect"] },
    { key: "global_economic_update", label: "Global Economic Update", upstream: ["collect:global_data_agent"] },
    { key: "global_economic_outlook", label: "Global Economic Outlook", upstream: ["collect:global_data_agent"] },
    { key: "nigeria_output_growth", label: "Nigeria's Output Growth", upstream: ["collect:nigeria_real_sector_agent"] },
    { key: "output_growth_outlook", label: "Output Growth Outlook", upstream: ["collect:nigeria_real_sector_agent"] },
    { key: "price_update", label: "Price Update", upstream: ["collect:price_cpi_agent"] },
    { key: "fiscal_operations", label: "Fiscal Operations Update", upstream: ["collect:fiscal_faac_agent"] },
    { key: "conclusion", label: "Conclusion", upstream: ["collect", "draft"] },
    { key: "country_in_focus", label: "Country in Focus", upstream: ["collect:country_focus_agent"] },
  ],
  chart: [
    { key: "chart_agent", label: "Chart / Figure Agent", upstream: ["collect", "draft"] },
  ],
  edit: [
    { key: "editor_agent", label: "Editor / Consistency Agent", upstream: ["collect", "draft"] },
  ],
};

// Sections with a real, filled-in prompt template — see prompts/sections/.
const READY_DRAFT_SECTIONS = new Set(["price_update"]);

const REPORT_ORDER = AGENTS.draft.map((a) => a.key);

let state = loadState();
let currentPhase = "collect";
let currentAgentKey = AGENTS.collect[0].key;
let schema = null;
const promptTextCache = new Map();
const editionCache = new Map();

// --- state persistence ---------------------------------------------------------

function emptyState() {
  const s = {};
  for (const phase of Object.keys(AGENTS)) {
    s[phase] = {};
    for (const agent of AGENTS[phase]) {
      s[phase][agent.key] = { input: "", output: "" };
    }
  }
  return s;
}

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return emptyState();
    const parsed = JSON.parse(raw);
    const fresh = emptyState();
    for (const phase of Object.keys(fresh)) {
      for (const key of Object.keys(fresh[phase])) {
        if (parsed[phase] && parsed[phase][key]) fresh[phase][key] = parsed[phase][key];
      }
    }
    return fresh;
  } catch {
    return emptyState();
  }
}

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function slot(phase, key) {
  return state[phase][key];
}

// --- helpers ---------------------------------------------------------------------

function findAgent(phase, key) {
  return AGENTS[phase].find((a) => a.key === key);
}

function allAgentsFlat() {
  return Object.keys(AGENTS).flatMap((phase) => AGENTS[phase].map((a) => ({ phase, ...a })));
}

function totalSlots() {
  return allAgentsFlat().length;
}

function completedSlots() {
  return allAgentsFlat().filter((a) => slot(a.phase, a.key).output.trim()).length;
}

async function fetchText(url) {
  if (!promptTextCache.has(url)) {
    const res = await fetch(url);
    if (!res.ok) throw new Error(`${url}: ${res.status}`);
    promptTextCache.set(url, await res.text());
  }
  return promptTextCache.get(url);
}

async function getSchema() {
  if (!schema) schema = await fetch(`${SCHEMAS_DIR}/section_data_schema.json`).then((r) => r.json());
  return schema;
}

async function getEditionRecord(id) {
  if (!editionCache.has(id)) {
    editionCache.set(id, await fetch(`${DATA_DIR}/processed/${id}.json`).then((r) => r.json()));
  }
  return editionCache.get(id);
}

function combinedJSON(agentRefs) {
  // agentRefs: array of "phase" (all agents in phase) or "phase:key" (one agent).
  const out = {};
  for (const ref of agentRefs) {
    const [phase, key] = ref.includes(":") ? ref.split(":") : [ref, null];
    const agents = key ? [findAgent(phase, key)] : AGENTS[phase];
    for (const agent of agents) {
      const val = slot(phase, agent.key).output.trim();
      if (!val) continue;
      try {
        out[agent.key] = JSON.parse(val);
      } catch {
        out[agent.key] = val;
      }
    }
  }
  return out;
}

// --- file parsing (client-side only — nothing leaves the browser) --------------

async function parseFile(file) {
  const name = file.name.toLowerCase();
  if (name.endsWith(".csv") || name.endsWith(".txt") || name.endsWith(".md") || name.endsWith(".json")) {
    return file.text();
  }
  if (name.endsWith(".xlsx") || name.endsWith(".xls")) {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    return wb.SheetNames.map((sn) => `# Sheet: ${sn}\n${XLSX.utils.sheet_to_csv(wb.Sheets[sn])}`).join("\n\n");
  }
  if (name.endsWith(".docx")) {
    const buf = await file.arrayBuffer();
    const result = await mammoth.extractRawText({ arrayBuffer: buf });
    return result.value;
  }
  if (name.endsWith(".pdf")) {
    const buf = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
    let text = "";
    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      const content = await page.getTextContent();
      text += content.items.map((it) => it.str).join(" ") + "\n\n";
    }
    return text;
  }
  throw new Error(`Unsupported file type: ${file.name}`);
}

// --- prompt building ---------------------------------------------------------------

function schemaPropertiesFor(key) {
  const props = schema.properties[key]?.properties || {};
  const shape = {};
  for (const [field, def] of Object.entries(props)) {
    shape[field] = def.description ? `${def.type} — ${def.description}` : def.type;
  }
  return shape;
}

async function buildCollectPrompt(agent, inputText) {
  await getSchema();
  const shape = schemaPropertiesFor(agent.schemaKey);
  const system =
    `You are the ${agent.label} for CAPE's monthly Economic Performance and Prospect Bulletin.\n\n` +
    `Extract ONLY the following fields from the source text the user provides, as strict JSON matching this shape ` +
    `(use null for anything not present — never invent, estimate, or carry over a figure that isn't in the source):\n\n` +
    `${JSON.stringify(shape, null, 2)}\n\n` +
    `Rules:\n` +
    `- Output JSON only. No prose, no markdown code fences.\n` +
    `- Every number must be traceable to the source text below.\n` +
    `- If a field isn't mentioned in the source, use null rather than guessing.\n` +
    `- Typical source: ${agent.sourceHint}.`;
  const user = `Source text:\n\n${inputText}`;
  return { system, user };
}

async function buildDraftPrompt(agent, inputText) {
  const sectionMd = await fetchText(`${PROMPTS_DIR}/sections/${agent.key}.md`).catch(() => "(no prompt template file found)");
  const styleGuideMd = await fetchText(`${PROMPTS_DIR}/style_guide.md`);
  const result = {
    system: `${sectionMd}\n\n---\n\n${styleGuideMd}`,
    user: `Draft this section from the following structured data. Do not state any figure that is not present in this JSON:\n\n${inputText}`,
  };
  if (!READY_DRAFT_SECTIONS.has(agent.key)) {
    result.warning = `prompts/sections/${agent.key}.md is still a Phase 1 placeholder — the system prompt below is incomplete.`;
  }
  return result;
}

function buildChartPrompt(inputText) {
  const system =
    "You are the Chart/Figure Agent for CAPE's monthly bulletin. Given the structured data and/or drafted " +
    "section text below, specify each of Figures 1-7 precisely: title, chart type, series/fields plotted, " +
    "time range, and axis labels — sourced only from the data provided, so the chart and the prose can never " +
    "disagree. Output as a numbered list, one figure per item, with the exact field names used for each series.";
  const user = `Structured data / drafted sections:\n\n${inputText}`;
  return { system, user };
}

function buildEditorPrompt(inputText) {
  const system =
    "You are the Editor/Consistency Agent for CAPE's monthly bulletin. Cross-check the drafted sections below " +
    "against the structured source data: flag any figure in the prose not traceable to the source JSON, check " +
    'MoM/YoY consistency, and confirm the Highlights section reflects the body. Return strict JSON: ' +
    '{"flags": [...], "diff_vs_last_month": "..."}. List unresolved/unverified claims explicitly in "flags".';
  const user = `Drafted sections + source data:\n\n${inputText}`;
  return { system, user };
}

async function buildPromptFor(phase, key, inputText) {
  const agent = findAgent(phase, key);
  if (phase === "collect") return buildCollectPrompt(agent, inputText);
  if (phase === "draft") return buildDraftPrompt(agent, inputText);
  if (phase === "chart") return buildChartPrompt(inputText);
  if (phase === "edit") return buildEditorPrompt(inputText);
  throw new Error(`Unknown phase ${phase}`);
}

// --- UI: pipeline nav ---------------------------------------------------------------

function renderProgress() {
  const pct = Math.round((completedSlots() / totalSlots()) * 100);
  document.getElementById("progress-fill").style.width = `${pct}%`;
}

function statusFor(phase, key) {
  const s = slot(phase, key);
  if (s.output.trim()) return "done";
  if (s.input.trim()) return "in-progress";
  return "empty";
}

function renderPhaseTabs() {
  const nav = document.getElementById("phase-tabs");
  nav.innerHTML = "";
  PHASES.forEach((phase) => {
    const done = AGENTS[phase.key].filter((a) => slot(phase.key, a.key).output.trim()).length;
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = `phase-tab phase-${phase.key}` + (phase.key === currentPhase ? " active" : "");
    btn.innerHTML = `<span class="phase-dot"></span>${phase.label} <span class="phase-count">${done}/${AGENTS[phase.key].length}</span>`;
    btn.addEventListener("click", () => {
      currentPhase = phase.key;
      currentAgentKey = AGENTS[phase.key][0].key;
      renderAll();
    });
    nav.appendChild(btn);
  });

  const reportBtn = document.createElement("button");
  reportBtn.type = "button";
  reportBtn.className = "phase-tab phase-report" + (currentPhase === "report" ? " active" : "");
  reportBtn.innerHTML = `<span class="phase-dot"></span>Report`;
  reportBtn.addEventListener("click", () => {
    currentPhase = "report";
    renderAll();
  });
  nav.appendChild(reportBtn);
}

function renderAgentChips() {
  const row = document.getElementById("agent-chips");
  row.innerHTML = "";
  if (currentPhase === "report") {
    row.hidden = true;
    return;
  }
  row.hidden = false;
  AGENTS[currentPhase].forEach((agent) => {
    const chip = document.createElement("button");
    chip.type = "button";
    chip.className = `agent-chip status-${statusFor(currentPhase, agent.key)}` + (agent.key === currentAgentKey ? " active" : "");
    chip.textContent = agent.label;
    chip.addEventListener("click", () => {
      currentAgentKey = agent.key;
      renderWorkbench();
      renderAgentChips();
    });
    row.appendChild(chip);
  });
}

// --- UI: workbench -------------------------------------------------------------------

function reuseOptionsFor(excludePhase, excludeKey) {
  const groups = [];
  const combineCollectHasAny = AGENTS.collect.some((a) => slot("collect", a.key).output.trim());
  const combineDraftHasAny = AGENTS.draft.some((a) => slot("draft", a.key).output.trim());
  groups.push({
    label: "Combined",
    options: [
      { value: "combine:collect", text: "All Collect agent outputs (combined JSON)", disabled: !combineCollectHasAny },
      { value: "combine:draft", text: "All Draft section outputs (combined JSON)", disabled: !combineDraftHasAny },
    ],
  });
  PHASES.forEach((phase) => {
    groups.push({
      label: phase.label,
      options: AGENTS[phase.key].map((a) => ({
        value: `${phase.key}:${a.key}`,
        text: a.label,
        disabled: !slot(phase.key, a.key).output.trim() || (phase.key === excludePhase && a.key === excludeKey),
      })),
    });
  });
  return groups;
}

function renderReuseSelect() {
  const select = document.getElementById("reuse-select");
  select.innerHTML = "";
  reuseOptionsFor(currentPhase, currentAgentKey).forEach((group) => {
    const optgroup = document.createElement("optgroup");
    optgroup.label = group.label;
    group.options.forEach((opt) => {
      const el = document.createElement("option");
      el.value = opt.value;
      el.textContent = opt.text;
      el.disabled = opt.disabled;
      optgroup.appendChild(el);
    });
    select.appendChild(optgroup);
  });
}

async function loadSampleOptions() {
  const select = document.getElementById("sample-select");
  select.innerHTML = '<option value="">— none —</option>';
  try {
    const editions = await fetch(`${DATA_DIR}/editions.json`).then((r) => r.json());
    editions.forEach((ed) => {
      const opt = document.createElement("option");
      opt.value = ed.id;
      opt.textContent = `${ed.month} ${ed.year}`;
      select.appendChild(opt);
    });
  } catch {
    // sample data optional; ignore if missing
  }
}

function renderWorkbench() {
  const agent = findAgent(currentPhase, currentAgentKey);
  const s = slot(currentPhase, currentAgentKey);

  document.getElementById("agent-title").textContent = agent.label;
  document.getElementById("agent-hint").textContent =
    currentPhase === "collect"
      ? `Typical source: ${agent.sourceHint}`
      : `Consumes: ${(agent.upstream || []).map((u) => (u.includes(":") ? findAgent(...u.split(":")).label : PHASES.find((p) => p.key === u).label)).join(", ") || "—"}`;

  const badge = document.getElementById("agent-status");
  const st = statusFor(currentPhase, currentAgentKey);
  badge.textContent = st === "done" ? "output saved" : st === "in-progress" ? "input in progress" : "not started";
  badge.className = `status-badge status-${st}`;

  const upstreamHint = document.getElementById("upstream-hint");
  if (currentPhase === "draft" && !READY_DRAFT_SECTIONS.has(agent.key)) {
    upstreamHint.hidden = false;
    upstreamHint.textContent = `⚠ prompts/sections/${agent.key}.md is still a Phase 1 placeholder — the built prompt will be incomplete until it's filled in.`;
  } else {
    upstreamHint.hidden = true;
  }

  document.getElementById("working-input").value = s.input;
  document.getElementById("agent-output").value = s.output;
  document.getElementById("parse-status").textContent = "";
  document.getElementById("prompt-out").hidden = true;
  renderReuseSelect();

  const sampleTab = document.querySelector('.input-tab[data-tab="sample"]');
  sampleTab.hidden = currentPhase !== "collect";
  if (sampleTab.hidden && document.querySelector('.input-tab.active').dataset.tab === "sample") {
    setActiveTab("upload");
  }
}

function setActiveTab(tab) {
  document.querySelectorAll(".input-tab").forEach((btn) => btn.classList.toggle("active", btn.dataset.tab === tab));
  document.getElementById("pane-upload").hidden = tab !== "upload";
  document.getElementById("pane-reuse").hidden = tab !== "reuse";
  document.getElementById("pane-paste").hidden = tab !== "paste";
  document.getElementById("pane-sample").hidden = tab !== "sample";
}

// --- UI: report assembly --------------------------------------------------------------

function renderReport() {
  document.getElementById("workbench-panel").hidden = currentPhase === "report";
  const reportPanel = document.getElementById("report-panel");
  reportPanel.hidden = false;

  const checklist = document.getElementById("report-checklist");
  checklist.innerHTML = "";
  REPORT_ORDER.forEach((key) => {
    const agent = findAgent("draft", key);
    const done = !!slot("draft", key).output.trim();
    const li = document.createElement("li");
    li.className = done ? "done" : "pending";
    li.textContent = agent.label;
    checklist.appendChild(li);
  });

  const assembly = document.getElementById("report-assembly");
  const parts = REPORT_ORDER.map((key) => {
    const agent = findAgent("draft", key);
    const output = slot("draft", key).output.trim();
    return `## ${agent.label}\n\n${output || "*(not drafted yet)*"}`;
  });
  assembly.textContent = parts.join("\n\n");
}

// --- top-level render ------------------------------------------------------------------

function renderAll() {
  renderProgress();
  renderPhaseTabs();
  renderAgentChips();
  if (currentPhase === "report") {
    renderReport();
  } else {
    document.getElementById("workbench-panel").hidden = false;
    document.getElementById("report-panel").hidden = true;
    renderWorkbench();
  }
}

// --- event wiring ------------------------------------------------------------------------

document.querySelectorAll(".input-tab").forEach((btn) => {
  btn.addEventListener("click", () => setActiveTab(btn.dataset.tab));
});

const dropzone = document.getElementById("dropzone");
const fileInput = document.getElementById("file-input");
dropzone.addEventListener("click", () => fileInput.click());
dropzone.addEventListener("dragover", (evt) => {
  evt.preventDefault();
  dropzone.classList.add("dragover");
});
dropzone.addEventListener("dragleave", () => dropzone.classList.remove("dragover"));
dropzone.addEventListener("drop", (evt) => {
  evt.preventDefault();
  dropzone.classList.remove("dragover");
  if (evt.dataTransfer.files.length) handleFile(evt.dataTransfer.files[0]);
});
fileInput.addEventListener("change", () => {
  if (fileInput.files.length) handleFile(fileInput.files[0]);
});

async function handleFile(file) {
  const statusEl = document.getElementById("parse-status");
  statusEl.textContent = `Parsing ${file.name}…`;
  try {
    const text = await parseFile(file);
    document.getElementById("working-input").value = text;
    slot(currentPhase, currentAgentKey).input = text;
    saveState();
    statusEl.textContent = `Parsed ${file.name} (${text.length.toLocaleString()} characters).`;
    renderAgentChips();
    renderPhaseTabs();
    renderProgress();
  } catch (err) {
    statusEl.textContent = `Could not parse ${file.name}: ${err.message}`;
  }
}

document.getElementById("reuse-load").addEventListener("click", () => {
  const value = document.getElementById("reuse-select").value;
  if (!value) return;
  let text = "";
  if (value.startsWith("combine:")) {
    const phase = value.split(":")[1];
    text = JSON.stringify(combinedJSON([phase]), null, 2);
  } else {
    const [phase, key] = value.split(":");
    text = slot(phase, key).output;
  }
  document.getElementById("working-input").value = text;
  slot(currentPhase, currentAgentKey).input = text;
  saveState();
  renderAgentChips();
  renderPhaseTabs();
  renderProgress();
});

document.getElementById("sample-select").addEventListener("change", async (evt) => {
  const id = evt.target.value;
  if (!id) return;
  const agent = findAgent(currentPhase, currentAgentKey);
  const record = await getEditionRecord(id);
  const text = JSON.stringify(record[agent.schemaKey] || {}, null, 2);
  document.getElementById("working-input").value = text;
  slot(currentPhase, currentAgentKey).input = text;
  saveState();
  renderAgentChips();
  renderPhaseTabs();
  renderProgress();
});

document.getElementById("working-input").addEventListener("input", (evt) => {
  slot(currentPhase, currentAgentKey).input = evt.target.value;
  saveState();
  renderAgentChips();
  renderPhaseTabs();
  renderProgress();
});

document.getElementById("build-prompt").addEventListener("click", async () => {
  const text = document.getElementById("working-input").value.trim();
  const out = document.getElementById("prompt-out");
  if (!text) {
    out.hidden = false;
    document.getElementById("p-system").value = "";
    document.getElementById("p-user").value = "Add a working input first (upload a file, reuse an agent's output, or paste text).";
    return;
  }
  try {
    const { system, user, warning } = await buildPromptFor(currentPhase, currentAgentKey, text);
    document.getElementById("p-system").value = warning ? `⚠ ${warning}\n\n${system}` : system;
    document.getElementById("p-user").value = user;
    out.hidden = false;
  } catch (err) {
    out.hidden = false;
    document.getElementById("p-system").value = `Could not build prompt: ${err.message}`;
    document.getElementById("p-user").value = "";
  }
});

document.querySelectorAll(".copy-button").forEach((btn) => {
  btn.addEventListener("click", async () => {
    const target = document.getElementById(btn.dataset.copyTarget);
    const ok = await copyText(target.value);
    btn.textContent = ok ? "Copied" : "Copy failed";
    setTimeout(() => (btn.textContent = "Copy"), 1200);
  });
});

document.getElementById("copy-both").addEventListener("click", async (evt) => {
  const system = document.getElementById("p-system").value;
  const user = document.getElementById("p-user").value;
  const ok = await copyText(`SYSTEM:\n${system}\n\nUSER:\n${user}`);
  evt.target.textContent = ok ? "Copied" : "Copy failed";
  setTimeout(() => (evt.target.textContent = "Copy both (ready to paste into claude.ai)"), 1200);
});

document.getElementById("agent-output").addEventListener("input", (evt) => {
  slot(currentPhase, currentAgentKey).output = evt.target.value;
  saveState();
  renderAgentChips();
  renderPhaseTabs();
  renderProgress();
});

document.getElementById("clear-output").addEventListener("click", () => {
  slot(currentPhase, currentAgentKey).output = "";
  document.getElementById("agent-output").value = "";
  saveState();
  renderAgentChips();
  renderPhaseTabs();
  renderProgress();
});

document.getElementById("copy-report").addEventListener("click", async (evt) => {
  const ok = await copyText(document.getElementById("report-assembly").textContent);
  evt.target.textContent = ok ? "Copied" : "Copy failed";
  setTimeout(() => (evt.target.textContent = "Copy full report"), 1200);
});

document.getElementById("session-export").addEventListener("click", () => {
  const blob = new Blob([JSON.stringify(state, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `cape-workbench-session-${new Date().toISOString().slice(0, 10)}.json`;
  a.click();
  URL.revokeObjectURL(url);
});

document.getElementById("session-import").addEventListener("change", async (evt) => {
  const file = evt.target.files[0];
  if (!file) return;
  try {
    const parsed = JSON.parse(await file.text());
    const fresh = emptyState();
    for (const phase of Object.keys(fresh)) {
      for (const key of Object.keys(fresh[phase])) {
        if (parsed[phase] && parsed[phase][key]) fresh[phase][key] = parsed[phase][key];
      }
    }
    state = fresh;
    saveState();
    renderAll();
  } catch (err) {
    alert(`Could not import session: ${err.message}`);
  }
  evt.target.value = "";
});

document.getElementById("session-reset").addEventListener("click", () => {
  if (!confirm("Clear all agent inputs and outputs in this browser? This can't be undone.")) return;
  state = emptyState();
  saveState();
  renderAll();
});

async function copyText(text) {
  try {
    await navigator.clipboard.writeText(text);
    return true;
  } catch {
    return false;
  }
}

async function init() {
  if (typeof pdfjsLib !== "undefined") {
    pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js";
  }
  await loadSampleOptions();
  renderAll();
}

init();
