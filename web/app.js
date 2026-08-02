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
    opt.disabled = !ready.includes(name);
    select.appendChild(opt);
  });
}

document.getElementById("generate-form").addEventListener("submit", async (evt) => {
  evt.preventDefault();
  const edition_id = document.getElementById("gen-edition").value;
  const section = document.getElementById("gen-section").value;
  const result = document.getElementById("generate-result");
  result.innerHTML = '<p class="loading">Calling the writer agent…</p>';

  try {
    const res = await fetch(`${API}/generate`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ edition_id, section }),
    });
    const data = await res.json();
    if (!res.ok) {
      result.innerHTML = `<p class="error">${data.detail || "Request failed."}</p>`;
      return;
    }
    result.textContent = data.text;
  } catch (err) {
    result.innerHTML = `<p class="error">${err.message}</p>`;
  }
});

loadTrend();
loadEditions();
loadSections();
