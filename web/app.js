const API = "/api";

const SERIES_COLORS = [
  "var(--series-1)",
  "var(--series-2)",
  "var(--series-3)",
];

// --- minimal SVG line chart, no external library ---------------------------
// series: [{ name, points: [{ x: "2026-01", y: 53.2 }, ...] }]
function renderLineChart(container, series, { yFormat = (v) => v } = {}) {
  container.innerHTML = "";
  const width = 520;
  const height = 220;
  const margin = { top: 12, right: 16, bottom: 26, left: 40 };
  const plotW = width - margin.left - margin.right;
  const plotH = height - margin.top - margin.bottom;

  const months = [...new Set(series.flatMap((s) => s.points.map((p) => p.x)))].sort();
  const xIndex = new Map(months.map((m, i) => [m, i]));
  const xScale = (m) => margin.left + (months.length <= 1 ? plotW / 2 : (xIndex.get(m) / (months.length - 1)) * plotW);

  const values = series.flatMap((s) => s.points.map((p) => p.y));
  const yMin = Math.min(...values);
  const yMax = Math.max(...values);
  const pad = (yMax - yMin) * 0.15 || Math.abs(yMax) * 0.1 || 1;
  const yLo = yMin - pad;
  const yHi = yMax + pad;
  const yScale = (v) => margin.top + plotH - ((v - yLo) / (yHi - yLo)) * plotH;

  const svgNS = "http://www.w3.org/2000/svg";
  const svg = document.createElementNS(svgNS, "svg");
  svg.setAttribute("viewBox", `0 0 ${width} ${height}`);
  svg.setAttribute("role", "img");
  svg.setAttribute("aria-label", series.map((s) => s.name).join(", ") + " trend chart");

  // gridlines (y)
  const yTicks = 4;
  for (let i = 0; i <= yTicks; i++) {
    const v = yLo + ((yHi - yLo) * i) / yTicks;
    const y = yScale(v);
    const line = document.createElementNS(svgNS, "line");
    line.setAttribute("x1", margin.left);
    line.setAttribute("x2", width - margin.right);
    line.setAttribute("y1", y);
    line.setAttribute("y2", y);
    line.setAttribute("stroke", "var(--gridline)");
    line.setAttribute("stroke-width", "1");
    svg.appendChild(line);

    const label = document.createElementNS(svgNS, "text");
    label.setAttribute("x", margin.left - 6);
    label.setAttribute("y", y + 3);
    label.setAttribute("text-anchor", "end");
    label.setAttribute("font-size", "9");
    label.setAttribute("fill", "var(--text-muted)");
    label.textContent = yFormat(v);
    svg.appendChild(label);
  }

  // x labels (thin out if many months)
  const step = Math.ceil(months.length / 6) || 1;
  months.forEach((m, i) => {
    if (i % step !== 0 && i !== months.length - 1) return;
    const label = document.createElementNS(svgNS, "text");
    label.setAttribute("x", xScale(m));
    label.setAttribute("y", height - 6);
    label.setAttribute("text-anchor", "middle");
    label.setAttribute("font-size", "9");
    label.setAttribute("fill", "var(--text-muted)");
    label.textContent = m.slice(2); // "26-03"
    svg.appendChild(label);
  });

  const tooltip = document.createElement("div");
  tooltip.className = "tooltip";
  container.style.position = "relative";

  series.forEach((s, si) => {
    const color = SERIES_COLORS[si % SERIES_COLORS.length];
    const pts = s.points.filter((p) => p.y !== null && p.y !== undefined);
    if (pts.length > 1) {
      const d = pts
        .map((p, i) => `${i === 0 ? "M" : "L"} ${xScale(p.x)} ${yScale(p.y)}`)
        .join(" ");
      const path = document.createElementNS(svgNS, "path");
      path.setAttribute("d", d);
      path.setAttribute("fill", "none");
      path.setAttribute("stroke", color);
      path.setAttribute("stroke-width", "2");
      path.setAttribute("stroke-linecap", "round");
      svg.appendChild(path);
    }
    pts.forEach((p) => {
      const hit = document.createElementNS(svgNS, "circle");
      hit.setAttribute("cx", xScale(p.x));
      hit.setAttribute("cy", yScale(p.y));
      hit.setAttribute("r", "9");
      hit.setAttribute("fill", "transparent");
      hit.style.cursor = "pointer";
      hit.addEventListener("mousemove", (evt) => {
        const rect = container.getBoundingClientRect();
        tooltip.style.left = `${evt.clientX - rect.left + 10}px`;
        tooltip.style.top = `${evt.clientY - rect.top - 10}px`;
        tooltip.textContent = `${s.name} · ${p.x}: ${yFormat(p.y)}`;
        tooltip.style.opacity = "1";
      });
      hit.addEventListener("mouseleave", () => {
        tooltip.style.opacity = "0";
      });
      svg.appendChild(hit);

      const dot = document.createElementNS(svgNS, "circle");
      dot.setAttribute("cx", xScale(p.x));
      dot.setAttribute("cy", yScale(p.y));
      dot.setAttribute("r", "3");
      dot.setAttribute("fill", color);
      dot.setAttribute("pointer-events", "none");
      svg.appendChild(dot);
    });
  });

  container.appendChild(svg);
  container.appendChild(tooltip);

  if (series.length > 1) {
    const legend = document.createElement("div");
    legend.className = "legend";
    series.forEach((s, si) => {
      const item = document.createElement("span");
      item.className = "legend-item";
      const swatch = document.createElement("span");
      swatch.className = "legend-swatch";
      swatch.style.background = SERIES_COLORS[si % SERIES_COLORS.length];
      item.appendChild(swatch);
      item.appendChild(document.createTextNode(s.name));
      legend.appendChild(item);
    });
    container.appendChild(legend);
  }
}

// --- data loading ------------------------------------------------------------

async function loadTrend() {
  const trend = await fetch(`${API}/trend`).then((r) => r.json());

  renderLineChart(document.getElementById("chart-pmi"), [
    { name: "Stanbic IBTC", points: trend.pmi_stanbic.map((d) => ({ x: d.month, y: d.value })) },
    { name: "CBN Composite", points: trend.pmi_cbn_composite.map((d) => ({ x: d.month, y: d.value })) },
  ], { yFormat: (v) => v.toFixed(1) });

  renderLineChart(document.getElementById("chart-inflation"), [
    { name: "Headline", points: trend.headline_inflation_yoy.map((d) => ({ x: d.month, y: d.value })) },
    { name: "Food", points: trend.food_inflation_yoy.map((d) => ({ x: d.month, y: d.value })) },
    { name: "Core", points: trend.core_inflation_yoy.map((d) => ({ x: d.month, y: d.value })) },
  ], { yFormat: (v) => `${v.toFixed(1)}%` });

  renderLineChart(document.getElementById("chart-faac"), [
    { name: "FAAC total", points: trend.faac_total_allocation_ngn_bn.map((d) => ({ x: d.month, y: d.value })) },
  ], { yFormat: (v) => v.toFixed(0) });

  renderLineChart(document.getElementById("chart-epu"), [
    { name: "EPU index", points: trend.epu_index.map((d) => ({ x: d.month, y: d.value })) },
  ], { yFormat: (v) => v.toFixed(0) });
}

async function loadEditions() {
  const editions = await fetch(`${API}/editions`).then((r) => r.json());
  const list = document.getElementById("edition-list");
  const genSelect = document.getElementById("gen-edition");
  list.innerHTML = "";
  genSelect.innerHTML = "";

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
    genSelect.appendChild(opt);
  });

  if (editions.length) selectEdition(editions[editions.length - 1].id, list);
}

async function selectEdition(id, listEl) {
  [...listEl.children].forEach((li) => li.classList.toggle("active", li.dataset.id === id));
  const record = await fetch(`${API}/editions/${id}`).then((r) => r.json());
  const detail = document.getElementById("edition-detail");

  const renderBlock = (title, obj) => {
    if (!obj) return "";
    const rows = Object.entries(obj)
      .filter(([k]) => k !== "notes")
      .map(([k, v]) => `<dt>${k}</dt><dd>${v === null ? "—" : v}</dd>`)
      .join("");
    const notes = obj.notes ? `<p class="notes">${obj.notes}</p>` : "";
    return `<h4>${title}</h4><dl>${rows}</dl>${notes}`;
  };

  detail.innerHTML = [
    renderBlock("Nigeria output", record.nigeria_output),
    renderBlock("Prices", record.prices),
    renderBlock("Fiscal", record.fiscal),
    renderBlock("Country in focus", record.country_in_focus),
  ].join("");

  if (record.extraction_notes) {
    const p = document.createElement("p");
    p.className = "notes";
    p.textContent = record.extraction_notes;
    detail.appendChild(p);
  }
}

async function loadSections() {
  const { ready, all } = await fetch(`${API}/sections`).then((r) => r.json());
  const select = document.getElementById("gen-section");
  select.innerHTML = "";
  all.forEach((name) => {
    const opt = document.createElement("option");
    opt.value = name;
    opt.textContent = ready.includes(name) ? name.replace(/_/g, " ") : `${name.replace(/_/g, " ")} (not ready)`;
    select.appendChild(opt);
  });
}

// --- data/questions panel ----------------------------------------------------
// Mirrors SECTION_DATA_KEY in src/cape_newsletter/api.py: which key in a
// processed edition's JSON holds the structured data for each section.
const SECTION_DATA_KEY = {
  global_economic_update: "global",
  global_economic_outlook: "global",
  nigeria_output_growth: "nigeria_output",
  output_growth_outlook: "nigeria_output",
  price_update: "prices",
  fiscal_operations: "fiscal",
  country_in_focus: "country_in_focus",
};

// Fields worth calling out as manual editorial input, not pipeline-derived —
// see docs/architecture.md's human-in-the-loop notes.
const MANUAL_FIELD_HINTS = {
  country: "Editorial choice — supplied manually each month, not inferred from data.",
};

const editionRecordCache = new Map();
let savedSectionData = null;
const dataFieldInputs = new Map();

async function getEditionRecord(id) {
  if (!editionRecordCache.has(id)) {
    editionRecordCache.set(id, await fetch(`${API}/editions/${id}`).then((r) => r.json()));
  }
  return editionRecordCache.get(id);
}

function renderDataFields(data) {
  const container = document.getElementById("gen-data-fields");
  container.innerHTML = "";
  dataFieldInputs.clear();

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

    dataFieldInputs.set(key, input);
    container.appendChild(wrap);
  });
}

function collectDataFromFields() {
  const result = {};
  dataFieldInputs.forEach((input, key) => {
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
  const edition_id = document.getElementById("gen-edition").value;
  const section = document.getElementById("gen-section").value;
  const fieldset = document.getElementById("gen-data-fieldset");
  const dataKey = SECTION_DATA_KEY[section];

  if (!edition_id || !dataKey) {
    fieldset.hidden = true;
    savedSectionData = null;
    return;
  }

  const record = await getEditionRecord(edition_id);
  savedSectionData = record[dataKey] || {};
  fieldset.hidden = false;
  renderDataFields(savedSectionData);
}

document.getElementById("gen-edition").addEventListener("change", updateDataPanel);
document.getElementById("gen-section").addEventListener("change", updateDataPanel);
document.getElementById("gen-data-reset").addEventListener("click", () => {
  if (savedSectionData) renderDataFields(savedSectionData);
});

document.getElementById("generate-form").addEventListener("submit", async (evt) => {
  evt.preventDefault();
  const edition_id = document.getElementById("gen-edition").value;
  const section = document.getElementById("gen-section").value;
  const result = document.getElementById("generate-result");
  result.innerHTML = '<p class="loading">Calling the writer agent…</p>';

  const data = dataFieldInputs.size ? collectDataFromFields() : undefined;

  try {
    const res = await fetch(`${API}/generate`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ edition_id, section, data }),
    });
    const resData = await res.json();
    if (!res.ok) {
      result.innerHTML = `<p class="error">${resData.detail || "Request failed."}</p>`;
      return;
    }
    result.textContent = resData.text;
  } catch (err) {
    result.innerHTML = `<p class="error">${err.message}</p>`;
  }
});

async function init() {
  await loadTrend();
  await loadEditions();
  await loadSections();
  await updateDataPanel();
}

init();
