// Farming CBA Decision Aid — Newcastle Business School
// Full upgrade using the real faba beans trial dataset.
// - Automatically loads faba_beans_trial_clean_named.tsv on start.
// - Uses all rows and columns; no sampling or removal.
// - Detects key columns, including replicate_id, treatment_name, is_control,
//   yield_t_ha, total_cost_per_ha (variable cost proxy), and
//   cost_amendment_input_per_ha (capital cost proxy).
// - Treats the measured control as the baseline for all comparisons.
// - Presents a control-centric comparison table with plain-language labels.
// - Provides leaderboard, charts, exports (TSV/CSV/XLSX), and AI summary prompt.

(() => {
  "use strict";

  // =========================
  // 0) STATE
  // =========================

  const state = {
    rawText: "",
    headers: [],
    rows: [],
    columnMap: null,
    treatments: [], // aggregated per treatment
    controlName: null,
    params: {
      pricePerTonne: 500,
      years: 10,
      persistenceYears: 10,
      discountRate: 5
    },
    results: {
      treatments: [],
      control: null
    },
    charts: {
      netProfitDelta: null,
      costsBenefits: null
    }
  };

  // =========================
  // 1) UTILITIES
  // =========================

  function showToast(message, type = "info", timeoutMs = 3200) {
    const container = document.getElementById("toastContainer");
    if (!container) return;
    const toast = document.createElement("div");
    toast.className = `toast ${type}`;
    toast.innerHTML = `
      <div>${message}</div>
      <button class="toast-dismiss" aria-label="Dismiss">&times;</button>
    `;
    container.appendChild(toast);

    const remove = () => {
      if (!toast.parentElement) return;
      toast.style.opacity = "0";
      toast.style.transform = "translateY(4px)";
      setTimeout(() => {
        if (toast.parentElement) {
          toast.parentElement.removeChild(toast);
        }
      }, 150);
    };

    toast.querySelector(".toast-dismiss").addEventListener("click", remove);
    setTimeout(remove, timeoutMs);
  }

  function parseNumber(value) {
    if (value === null || value === undefined) return NaN;
    if (typeof value === "number") return Number.isFinite(value) ? value : NaN;
    const trimmed = String(value).trim();
    if (!trimmed) return NaN;
    const cleaned = trimmed.replace(/,/g, "");
    const num = Number(cleaned);
    return Number.isFinite(num) ? num : NaN;
  }

  function parseBoolean(value) {
    if (value === null || value === undefined) return false;
    const s = String(value).trim().toLowerCase();
    return s === "true" || s === "1" || s === "yes" || s === "y";
  }

  function discountFactorSum(discountRatePct, years) {
    const r = discountRatePct / 100;
    if (years <= 0) return 0;
    if (r === 0) return years;
    let sum = 0;
    for (let t = 1; t <= years; t++) {
      sum += 1 / Math.pow(1 + r, t);
    }
    return sum;
  }

  function meanIgnoringNaN(values) {
    if (!values || values.length === 0) return NaN;
    let sum = 0;
    let n = 0;
    for (const v of values) {
      const x = parseNumber(v);
      if (!Number.isNaN(x)) {
        sum += x;
        n++;
      }
    }
    return n === 0 ? NaN : sum / n;
  }

  function formatCurrency(value) {
    if (value === null || value === undefined || Number.isNaN(value)) return "–";
    try {
      return new Intl.NumberFormat("en-AU", {
        style: "currency",
        currency: "AUD",
        maximumFractionDigits: Math.abs(value) < 1000 ? 2 : 0
      }).format(value);
    } catch (_) {
      return value.toFixed(0);
    }
  }

  function formatNumber(value, decimals = 2) {
    if (value === null || value === undefined || Number.isNaN(value)) return "–";
    return value.toFixed(decimals);
  }

  function clone(obj) {
    return JSON.parse(JSON.stringify(obj));
  }

  // =========================
  // 2) DATA PARSING
  // =========================

  function detectDelimiter(headerLine) {
    if (!headerLine) return "\t";
    const tabCount = (headerLine.match(/\t/g) || []).length;
    const commaCount = (headerLine.match(/,/g) || []).length;
    return tabCount >= commaCount ? "\t" : ",";
  }

  function parseDelimitedText(text) {
    const trimmed = text.replace(/\r\n/g, "\n").trim();
    if (!trimmed) {
      throw new Error("No data found in the provided text.");
    }

    const lines = trimmed.split("\n").filter((l) => l.trim().length > 0);
    const delimiter = detectDelimiter(lines[0]);
    const headers = lines[0].split(delimiter).map((h) => h.trim());
    const rows = [];

    for (let i = 1; i < lines.length; i++) {
      const parts = lines[i].split(delimiter);
      const row = {};
      for (let j = 0; j < headers.length; j++) {
        row[headers[j]] = j < parts.length ? parts[j] : "";
      }
      rows.push(row);
    }
    return { headers, rows };
  }

  // Map required semantic columns to actual dataset headers
  function detectColumnMap(headers) {
    const lcHeaders = headers.map((h) => h.toLowerCase());

    const findCol = (aliases) => {
      for (const alias of aliases) {
        const idx = lcHeaders.indexOf(alias.toLowerCase());
        if (idx !== -1) return headers[idx];
      }
      return null;
    };

    const map = {
      plotId: findCol(["plot_id"]),
      replicate: findCol(["replicate_id", "replicate", "rep", "rep_no"]),
      treatmentName: findCol(["treatment_name", "amendment_name", "treatment"]),
      isControl: findCol(["is_control", "control_flag", "control"]),
      yield: findCol(["yield_t_ha", "yield", "yield_t/ha"]),
      variableCost: findCol([
        "total_cost_per_ha",
        "variable_cost_per_ha",
        "total_cost_per_ha_raw"
      ]),
      capitalCost: findCol([
        "capital_cost_per_ha",
        "cost_amendment_input_per_ha"
      ])
    };

    return map;
  }

  // =========================
  // 3) AGGREGATION & CBA
  // =========================

  function aggregateTreatments() {
    const { headers, rows } = state;
    if (!headers.length || !rows.length) return;

    const columnMap = detectColumnMap(headers);
    state.columnMap = columnMap;

    const summaryByName = new Map();
    let controlNameFromFlag = null;

    for (const row of rows) {
      const tNameRaw = columnMap.treatmentName
        ? row[columnMap.treatmentName]
        : "";
      const treatmentName = tNameRaw ? String(tNameRaw).trim() : "";
      if (!treatmentName) continue;

      if (!summaryByName.has(treatmentName)) {
        summaryByName.set(treatmentName, {
          name: treatmentName,
          replicates: [],
          yields: [],
          variableCosts: [],
          capitalCosts: [],
          isControlFlagged: false
        });
      }
      const t = summaryByName.get(treatmentName);

      const repVal = columnMap.replicate ? row[columnMap.replicate] : null;
      if (repVal !== null && repVal !== undefined && String(repVal).trim()) {
        t.replicates.push(repVal);
      }

      if (columnMap.yield) {
        const y = parseNumber(row[columnMap.yield]);
        if (!Number.isNaN(y)) t.yields.push(y);
      }
      if (columnMap.variableCost) {
        const vc = parseNumber(row[columnMap.variableCost]);
        if (!Number.isNaN(vc)) t.variableCosts.push(vc);
      }
      if (columnMap.capitalCost) {
        const cc = parseNumber(row[columnMap.capitalCost]);
        if (!Number.isNaN(cc)) t.capitalCosts.push(cc);
      }
      if (columnMap.isControl) {
        const isCtrl = parseBoolean(row[columnMap.isControl]);
        if (isCtrl) {
          t.isControlFlagged = true;
          controlNameFromFlag = treatmentName;
        }
      }
    }

    // Decide control treatment
    let controlName = controlNameFromFlag;
    if (!controlName) {
      for (const [name, t] of summaryByName.entries()) {
        if (name.toLowerCase().includes("control")) {
          controlName = name;
          break;
        }
      }
    }
    if (!controlName && summaryByName.size > 0) {
      controlName = summaryByName.keys().next().value;
    }
    state.controlName = controlName || null;

    const treatments = [];
    for (const [name, t] of summaryByName.entries()) {
      const avgYield = meanIgnoringNaN(t.yields);
      const avgVarCost = meanIgnoringNaN(t.variableCosts);
      const avgCapCost = Number.isNaN(meanIgnoringNaN(t.capitalCosts))
        ? 0
        : meanIgnoringNaN(t.capitalCosts);
      treatments.push({
        name,
        isControl: name === controlName,
        avgYield,
        avgVarCost,
        avgCapCost
      });
    }

    state.treatments = treatments;
  }

  function computeCBA() {
    const { treatments, controlName, params } = state;
    if (!treatments.length || !controlName) return;

    const control = treatments.find((t) => t.name === controlName) || treatments[0];

    const price = parseNumber(params.pricePerTonne) || 0;
    const years = Math.max(1, parseInt(params.years, 10) || 1);
    const persistenceYears = Math.max(
      1,
      Math.min(years, parseInt(params.persistenceYears, 10) || years)
    );
    const discountRate = parseNumber(params.discountRate);
    const factorBenefits = discountFactorSum(discountRate, persistenceYears);
    const factorCosts = discountFactorSum(discountRate, years);

    const results = [];
    for (const t of treatments) {
      const avgYield = parseNumber(t.avgYield);
      const avgVarCost = parseNumber(t.avgVarCost);
      const avgCapCost = parseNumber(t.avgCapCost);

      const pvBenefits = Number.isNaN(avgYield)
        ? NaN
        : avgYield * price * factorBenefits;
      const pvVarCosts = Number.isNaN(avgVarCost)
        ? NaN
        : avgVarCost * factorCosts;
      const pvCapCosts = Number.isNaN(avgCapCost) ? 0 : avgCapCost;
      const pvTotalCosts =
        Number.isNaN(pvVarCosts) && Number.isNaN(pvCapCosts)
          ? NaN
          : (Number.isNaN(pvVarCosts) ? 0 : pvVarCosts) + pvCapCosts;
      const npv =
        Number.isNaN(pvBenefits) || Number.isNaN(pvTotalCosts)
          ? NaN
          : pvBenefits - pvTotalCosts;
      const bcr =
        !Number.isNaN(pvBenefits) && pvTotalCosts > 0
          ? pvBenefits / pvTotalCosts
          : NaN;
      const roi =
        !Number.isNaN(npv) && pvTotalCosts > 0
          ? (npv / pvTotalCosts) * 100
          : NaN;

      results.push({
        name: t.name,
        isControl: t.name === controlName,
        avgYield,
        avgVarCost,
        avgCapCost,
        pvBenefits,
        pvTotalCosts,
        npv,
        bcr,
        roi
      });
    }

    // Control metrics
    const controlRes = results.find((r) => r.isControl) || results[0];
    const cPvBenefits = controlRes.pvBenefits;
    const cPvCosts = controlRes.pvTotalCosts;
    const cNpv = controlRes.npv;

    // Differences vs control and ranking
    for (const r of results) {
      r.deltaPvBenefits = r.pvBenefits - cPvBenefits;
      r.deltaPvCosts = r.pvTotalCosts - cPvCosts;
      r.deltaNpv = r.npv - cNpv;
    }

    results.sort((a, b) => {
      const aVal = Number.isNaN(a.npv) ? -Infinity : a.npv;
      const bVal = Number.isNaN(b.npv) ? -Infinity : b.npv;
      return bVal - aVal;
    });

    results.forEach((r, idx) => {
      r.rank = idx + 1;
    });

    state.results = {
      treatments: results,
      control: controlRes,
      price,
      years,
      persistenceYears,
      discountRate
    };
  }

  // =========================
  // 4) RENDERING
  // =========================

  function renderOverview() {
    const container = document.getElementById("overviewSummary");
    if (!container) return;
    container.innerHTML = "";

    if (!state.rows.length || !state.treatments.length) {
      container.innerHTML = `<p class="small muted">No dataset loaded yet.</p>`;
      return;
    }

    const nPlots = state.rows.length;
    const nTreatments = state.treatments.length;
    const nReplicates = new Set(
      state.rows
        .map((r) =>
          state.columnMap && state.columnMap.replicate
            ? r[state.columnMap.replicate]
            : null
        )
        .filter((x) => x !== null && x !== undefined && String(x).trim() !== "")
    ).size;

    const controlName = state.controlName || "Not identified";

    const cards = [
      {
        label: "Plots in dataset",
        value: nPlots
      },
      {
        label: "Distinct treatments",
        value: nTreatments
      },
      {
        label: "Replicates",
        value: nReplicates || "–"
      },
      {
        label: "Control treatment",
        value: controlName
      }
    ];

    for (const c of cards) {
      const div = document.createElement("div");
      div.className = "summary-card";
      div.innerHTML = `
        <div class="summary-card-label">${c.label}</div>
        <div class="summary-card-value">${c.value}</div>
      `;
      container.appendChild(div);
    }
  }

  function renderDataSummaryAndChecks() {
    const summaryEl = document.getElementById("dataSummary");
    const checksEl = document.getElementById("dataChecks");
    if (!summaryEl || !checksEl) return;

    if (!state.rows.length || !state.headers.length) {
      summaryEl.textContent = "No dataset loaded.";
      checksEl.innerHTML = "";
      return;
    }

    const nRows = state.rows.length;
    const nCols = state.headers.length;
    summaryEl.textContent = `${nRows} plot rows, ${nCols} columns. All rows and columns are used in calculations.`;

    const cm = state.columnMap;
    const checks = [];

    const addCheck = (ok, label, detail, severity) => {
      checks.push({ ok, label, detail, severity });
    };

    const addColCheck = (purpose, colName) => {
      if (colName) {
        addCheck(
          true,
          `${purpose} column found`,
          `"${colName}" is used for ${purpose.toLowerCase()}.`,
          "ok"
        );
      } else {
        addCheck(
          false,
          `${purpose} column missing`,
          `No column was found for ${purpose.toLowerCase()}. The tool will still run but some metrics may be incomplete.`,
          "warn"
        );
      }
    };

    addColCheck("Treatment name", cm.treatmentName);
    addColCheck("Control flag", cm.isControl);
    addColCheck("Replicate", cm.replicate);
    addColCheck("Yield", cm.yield);
    addColCheck("Variable cost", cm.variableCost);
    addColCheck("Capital cost", cm.capitalCost);

    const missingYieldRows = state.rows.filter((r) => {
      if (!cm.yield) return false;
      const v = r[cm.yield];
      return v === null || v === undefined || String(v).trim() === "";
    }).length;

    if (missingYieldRows > 0) {
      addCheck(
        false,
        "Missing yield values",
        `${missingYieldRows} rows have missing yield; these rows still count for costs but are excluded from yield averages.`,
        "warn"
      );
    } else {
      addCheck(
        true,
        "Yield values present",
        "All rows have yield values in the detected yield column.",
        "ok"
      );
    }

    checksEl.innerHTML = "";
    for (const c of checks) {
      const li = document.createElement("li");
      li.className = "check-item";
      const pillClass =
        c.severity === "err" ? "err" : c.severity === "warn" ? "warn" : "ok";
      li.innerHTML = `
        <span class="check-pill ${pillClass}">${c.ok ? "OK" : "Check"}</span>
        <span>${c.label}. ${c.detail}</span>
      `;
      checksEl.appendChild(li);
    }
  }

  function renderControlChoice() {
    const select = document.getElementById("controlChoice");
    if (!select) return;
    select.innerHTML = "";

    if (!state.treatments.length) return;

    for (const t of state.treatments) {
      const opt = document.createElement("option");
      opt.value = t.name;
      opt.textContent = t.name;
      if (t.name === state.controlName) {
        opt.selected = true;
      }
      select.appendChild(opt);
    }
  }

  function renderLeaderboard() {
    const container = document.getElementById("leaderboardContainer");
    const filterSelect = document.getElementById("leaderboardFilter");
    if (!container) return;

    container.innerHTML = "";

    const { treatments } = state.results;
    if (!treatments || !treatments.length) {
      container.innerHTML =
        '<p class="small muted">No results yet. Load data and apply scenario settings.</p>';
      return;
    }

    const filter = filterSelect ? filterSelect.value : "all";

    let rows = treatments.slice();
    const control = treatments.find((t) => t.isControl);

    if (filter === "topNPV") {
      rows = rows
        .slice()
        .filter((r) => !Number.isNaN(r.npv))
        .sort((a, b) => b.npv - a.npv)
        .slice(0, 5);
    } else if (filter === "topBCR") {
      rows = rows
        .slice()
        .filter((r) => !Number.isNaN(r.bcr))
        .sort((a, b) => b.bcr - a.bcr)
        .slice(0, 5);
    } else if (filter === "betterThanControl" && control) {
      rows = rows.filter((r) => r.npv > control.npv);
    }

    const table = document.createElement("table");
    table.className = "leaderboard-table";
    const thead = document.createElement("thead");
    thead.innerHTML = `
      <tr>
        <th>Rank</th>
        <th>Treatment</th>
        <th>Net profit over time</th>
        <th>Difference in net profit vs control</th>
      </tr>
    `;
    table.appendChild(thead);

    const tbody = document.createElement("tbody");
    for (const r of rows) {
      const tr = document.createElement("tr");
      if (r.isControl) {
        tr.classList.add("highlight-better");
      } else if (r.deltaNpv < 0) {
        tr.classList.add("highlight-worse");
      }

      const deltaText = r.isControl
        ? "Baseline"
        : `${r.deltaNpv >= 0 ? "+" : ""}${formatCurrency(r.deltaNpv)}`;

      tr.innerHTML = `
        <td>${r.rank}</td>
        <td>
          ${r.name}
          ${r.isControl ? '<span class="tag-control">Control</span>' : ""}
        </td>
        <td>${formatCurrency(r.npv)}</td>
        <td>${deltaText}</td>
      `;
      tbody.appendChild(tr);
    }
    table.appendChild(tbody);

    container.appendChild(table);
  }

  function renderComparisonTable() {
    const table = document.getElementById("comparisonTable");
    if (!table) return;

    table.innerHTML = "";

    const { treatments } = state.results;
    if (!treatments || !treatments.length) {
      table.innerHTML =
        '<tbody><tr><td class="small muted">No results yet. Load data and apply scenario settings.</td></tr></tbody>';
      return;
    }

    const control = treatments.find((t) => t.isControl) || treatments[0];
    const nonControl = treatments.filter((t) => !t.isControl);
    const ordered = [control, ...nonControl];

    const indicators = [
      { key: "pvBenefits", label: "Total benefits over time (discounted)" },
      { key: "pvTotalCosts", label: "Total costs over time (discounted)" },
      { key: "npv", label: "Net profit over time" },
      { key: "bcr", label: "Benefit per dollar spent" },
      { key: "roi", label: "Return on investment (percent)" },
      { key: "rank", label: "Overall ranking" },
      {
        key: "deltaNpv",
        label: "Difference in net profit compared with control"
      },
      {
        key: "deltaPvCosts",
        label: "Difference in total cost compared with control"
      }
    ];

    const thead = document.createElement("thead");
    const headRow = document.createElement("tr");
    headRow.innerHTML =
      '<th class="sticky-col">What is measured</th>' +
      ordered
        .map(
          (t) =>
            `<th>${t.isControl ? "Control (baseline)" : escapeHtml(t.name)}</th>`
        )
        .join("");
    thead.appendChild(headRow);
    table.appendChild(thead);

    const tbody = document.createElement("tbody");

    for (const ind of indicators) {
      const tr = document.createElement("tr");
      const firstCell = document.createElement("td");
      firstCell.className = "sticky-col";
      firstCell.innerHTML = `<span class="indicator-name">${ind.label}</span>`;
      tr.appendChild(firstCell);

      for (const t of ordered) {
        const r = treatments.find((x) => x.name === t.name);
        const td = document.createElement("td");
        let mainVal = "";
        let subVal = "";
        let subClass = "";

        if (ind.key === "pvBenefits") {
          mainVal = formatCurrency(r.pvBenefits);
          if (!r.isControl) {
            const delta = r.deltaPvBenefits;
            if (!Number.isNaN(delta)) {
              subVal = `${delta >= 0 ? "+" : ""}${formatCurrency(delta)} vs control`;
              subClass = delta >= 0 ? "cell-better" : "cell-worse";
            }
          }
        } else if (ind.key === "pvTotalCosts") {
          mainVal = formatCurrency(r.pvTotalCosts);
          if (!r.isControl) {
            const delta = r.deltaPvCosts;
            if (!Number.isNaN(delta)) {
              subVal = `${delta >= 0 ? "+" : ""}${formatCurrency(
                delta
              )} vs control`;
              subClass = delta <= 0 ? "cell-better" : "cell-worse";
            }
          }
        } else if (ind.key === "npv") {
          mainVal = formatCurrency(r.npv);
          if (!r.isControl) {
            const delta = r.deltaNpv;
            if (!Number.isNaN(delta)) {
              subVal = `${delta >= 0 ? "+" : ""}${formatCurrency(
                delta
              )} vs control`;
              subClass = delta >= 0 ? "cell-better" : "cell-worse";
            }
          } else {
            subVal = "Baseline level";
          }
        } else if (ind.key === "bcr") {
          mainVal = Number.isNaN(r.bcr) ? "–" : formatNumber(r.bcr, 2);
        } else if (ind.key === "roi") {
          mainVal = Number.isNaN(r.roi)
            ? "–"
            : `${formatNumber(r.roi, 1)}%`;
        } else if (ind.key === "rank") {
          mainVal = r.rank;
        } else if (ind.key === "deltaNpv") {
          mainVal = r.isControl
            ? "0 (baseline)"
            : `${r.deltaNpv >= 0 ? "+" : ""}${formatCurrency(r.deltaNpv)}`;
          if (!r.isControl) {
            subClass = r.deltaNpv >= 0 ? "cell-better" : "cell-worse";
          }
        } else if (ind.key === "deltaPvCosts") {
          mainVal = r.isControl
            ? "0 (baseline)"
            : `${r.deltaPvCosts >= 0 ? "+" : ""}${formatCurrency(
                r.deltaPvCosts
              )}`;
          if (!r.isControl) {
            subClass = r.deltaPvCosts <= 0 ? "cell-better" : "cell-worse";
          }
        }

        td.innerHTML = `
          <div class="cell-main">${mainVal}</div>
          <div class="cell-sub ${subClass}">${subVal}</div>
        `;
        tr.appendChild(td);
      }

      tbody.appendChild(tr);
    }

    table.appendChild(tbody);
  }

  // Escape HTML in treatment names
  function escapeHtml(str) {
    return String(str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  function renderCharts() {
    const { treatments } = state.results;
    if (!treatments || !treatments.length) return;

    const control = treatments.find((t) => t.isControl);
    const nonControl = treatments.filter((t) => !t.isControl);
    const all = treatments.slice();

    // Net profit vs control chart
    const ctxNet = document.getElementById("chartNetProfitVsControl");
    if (ctxNet) {
      if (state.charts.netProfitDelta) {
        state.charts.netProfitDelta.destroy();
      }
      const labels = nonControl.map((t) => t.name);
      const dataDelta = nonControl.map((t) => t.deltaNpv);

      state.charts.netProfitDelta = new Chart(ctxNet, {
        type: "bar",
        data: {
          labels,
          datasets: [
            {
              label: "Difference in net profit vs control (per hectare)",
              data: dataDelta
            }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          scales: {
            x: {
              ticks: { font: { size: 10 } }
            },
            y: {
              title: {
                display: true,
                text: "Difference in net profit (AUD per hectare)"
              }
            }
          },
          plugins: {
            legend: { display: false },
            tooltip: {
              callbacks: {
                label: (ctx) =>
                  `${ctx.dataset.label}: ${formatCurrency(ctx.parsed.y)}`
              }
            }
          }
        }
      });
    }

    // Costs and benefits chart
    const ctxCB = document.getElementById("chartCostsBenefits");
    if (ctxCB) {
      if (state.charts.costsBenefits) {
        state.charts.costsBenefits.destroy();
      }

      const labelsCB = all.map((t) =>
        t.isControl ? `${t.name} (control)` : t.name
      );
      const benefits = all.map((t) => t.pvBenefits);
      const costs = all.map((t) => t.pvTotalCosts);

      state.charts.costsBenefits = new Chart(ctxCB, {
        type: "bar",
        data: {
          labels: labelsCB,
          datasets: [
            {
              label: "Total benefits over time (discounted)",
              data: benefits
            },
            {
              label: "Total costs over time (discounted)",
              data: costs
            }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          scales: {
            x: {
              stacked: false,
              ticks: { font: { size: 10 } }
            },
            y: {
              stacked: false,
              title: {
                display: true,
                text: "AUD per hectare"
              }
            }
          },
          plugins: {
            tooltip: {
              callbacks: {
                label: (ctx) =>
                  `${ctx.dataset.label}: ${formatCurrency(ctx.parsed.y)}`
              }
            }
          }
        }
      });
    }
  }

  function buildAiBriefingPrompt() {
    const textarea = document.getElementById("aiBriefing");
    if (!textarea) return;

    const { treatments, control, price, years, persistenceYears, discountRate } =
      state.results;
    if (!treatments || !treatments.length || !control) {
      textarea.value =
        "No scenario is available yet. Load the trial dataset, apply scenario settings, and then regenerate this summary.";
      return;
    }

    const lines = [];
    lines.push(
      "You are reviewing a cost–benefit analysis based on a real faba beans soil amendment trial."
    );
    lines.push(
      "All figures are calculated from the full trial dataset without dropping any plots."
    );
    lines.push("");
    lines.push("Scenario settings:");
    lines.push(
      `- Grain price per tonne: ${formatCurrency(price)} per tonne of harvested grain.`
    );
    lines.push(
      `- Years of benefits and costs: ${years} years are included in the analysis.`
    );
    lines.push(
      `- Years that yield gains last: yield gains from treatments are assumed to last for ${persistenceYears} years.`
    );
    lines.push(
      `- Discount rate: ${discountRate.toFixed(
        1
      )} percent per year is used to express future flows in today’s dollars.`
    );
    lines.push("");
    lines.push(
      `The control treatment is "${control.name}". It is the baseline for all comparisons.`
    );
    lines.push(
      `Its discounted total benefits per hectare are ${formatCurrency(
        control.pvBenefits
      )}, its discounted total costs are ${formatCurrency(
        control.pvTotalCosts
      )}, and its net profit over time is ${formatCurrency(control.npv)}.`
    );
    lines.push("");
    lines.push(
      "For each other treatment, describe how it compares with the control in terms of total discounted benefits, total discounted costs, and net profit per hectare. Focus on practical interpretation for farmers and advisers."
    );
    lines.push("");
    lines.push("Key treatment results to interpret:");
    for (const t of treatments) {
      const label = t.isControl ? "Control" : "Treatment";
      const deltaNpvText = t.isControl
        ? "0 (baseline)"
        : `${t.deltaNpv >= 0 ? "higher" : "lower"} net profit of about ${formatCurrency(
            Math.abs(t.deltaNpv)
          )} per hectare compared with the control.`;
      const deltaCostText = t.isControl
        ? "0 (baseline)"
        : `${t.deltaPvCosts >= 0 ? "higher" : "lower"} discounted total costs by about ${formatCurrency(
            Math.abs(t.deltaPvCosts)
          )} per hectare relative to the control.`;

      lines.push(
        `- ${label} "${t.name}": net profit over time ${formatCurrency(
          t.npv
        )} per hectare; benefit per dollar spent around ${
          Number.isNaN(t.bcr) ? "not defined" : t.bcr.toFixed(2)
        }; return on investment about ${
          Number.isNaN(t.roi) ? "not defined" : t.roi.toFixed(1)
        } percent. This corresponds to ${deltaNpvText} It also has ${deltaCostText}`
      );
    }
    lines.push("");
    lines.push(
      "Write a concise narrative that explains which treatments look most promising, where gains come from higher yields versus lower costs, and how much better or worse they are than the control in practical terms."
    );
    lines.push(
      "Use clear language that farmers, agronomists, and policy makers can understand. Avoid technical shorthand such as NPV or BCR; instead, talk about total profit over time, total costs over time, and benefit per dollar spent."
    );

    textarea.value = lines.join("\n");
  }

  // =========================
  // 5) EXPORTS
  // =========================

  function exportCleanDatasetTSV() {
    if (!state.headers.length || !state.rows.length) {
      showToast("No dataset to export.", "error");
      return;
    }
    const headerLine = state.headers.join("\t");
    const lines = [headerLine];

    for (const row of state.rows) {
      const cells = state.headers.map((h) => {
        const v = row[h];
        return v === undefined || v === null ? "" : String(v);
      });
      lines.push(cells.join("\t"));
    }

    const blob = new Blob([lines.join("\n")], {
      type: "text/tab-separated-values;charset=utf-8;"
    });
    downloadBlob(blob, "faba_beans_dataset_clean.tsv");
    showToast("Cleaned dataset (TSV) downloaded.", "success");
  }

  function exportTreatmentSummaryCSV() {
    const { treatments } = state.results;
    if (!treatments || !treatments.length) {
      showToast("No treatment summary to export.", "error");
      return;
    }

    const headers = [
      "Treatment name",
      "Is control",
      "Average yield (t per ha)",
      "Average variable cost (per ha)",
      "Average capital cost (per ha)",
      "Total benefits over time (discounted)",
      "Total costs over time (discounted)",
      "Net profit over time",
      "Benefit per dollar spent",
      "Return on investment (percent)",
      "Rank",
      "Difference in net profit vs control",
      "Difference in total cost vs control"
    ];

    const lines = [];
    lines.push(headers.join(","));

    for (const t of treatments) {
      const row = [
        `"${t.name.replace(/"/g, '""')}"`,
        t.isControl ? "TRUE" : "FALSE",
        Number.isNaN(t.avgYield) ? "" : t.avgYield.toFixed(4),
        Number.isNaN(t.avgVarCost) ? "" : t.avgVarCost.toFixed(4),
        Number.isNaN(t.avgCapCost) ? "" : t.avgCapCost.toFixed(4),
        Number.isNaN(t.pvBenefits) ? "" : t.pvBenefits.toFixed(2),
        Number.isNaN(t.pvTotalCosts) ? "" : t.pvTotalCosts.toFixed(2),
        Number.isNaN(t.npv) ? "" : t.npv.toFixed(2),
        Number.isNaN(t.bcr) ? "" : t.bcr.toFixed(4),
        Number.isNaN(t.roi) ? "" : t.roi.toFixed(2),
        t.rank,
        Number.isNaN(t.deltaNpv) ? "" : t.deltaNpv.toFixed(2),
        Number.isNaN(t.deltaPvCosts) ? "" : t.deltaPvCosts.toFixed(2)
      ];
      lines.push(row.join(","));
    }

    const blob = new Blob([lines.join("\n")], {
      type: "text/csv;charset=utf-8;"
    });
    downloadBlob(blob, "treatment_summary.csv");
    showToast("Treatment summary (CSV) downloaded.", "success");
  }

  function exportComparisonCSV() {
    const { treatments } = state.results;
    if (!treatments || !treatments.length) {
      showToast("No results table to export.", "error");
      return;
    }

    const indicators = [
      { key: "pvBenefits", label: "Total benefits over time (discounted)" },
      { key: "pvTotalCosts", label: "Total costs over time (discounted)" },
      { key: "npv", label: "Net profit over time" },
      { key: "bcr", label: "Benefit per dollar spent" },
      { key: "roi", label: "Return on investment (percent)" },
      { key: "rank", label: "Overall ranking" },
      {
        key: "deltaNpv",
        label: "Difference in net profit compared with control"
      },
      {
        key: "deltaPvCosts",
        label: "Difference in total cost compared with control"
      }
    ];

    const control = treatments.find((t) => t.isControl) || treatments[0];
    const nonControl = treatments.filter((t) => !t.isControl);
    const ordered = [control, ...nonControl];

    const header = ["What is measured"];
    for (const t of ordered) {
      header.push(t.isControl ? "Control (baseline)" : t.name);
    }

    const lines = [];
    lines.push(header.map(csvEscape).join(","));

    for (const ind of indicators) {
      const row = [ind.label];
      for (const t of ordered) {
        const r = treatments.find((x) => x.name === t.name);
        let value = "";
        if (ind.key === "pvBenefits") {
          value = Number.isNaN(r.pvBenefits) ? "" : r.pvBenefits.toFixed(2);
        } else if (ind.key === "pvTotalCosts") {
          value = Number.isNaN(r.pvTotalCosts) ? "" : r.pvTotalCosts.toFixed(2);
        } else if (ind.key === "npv") {
          value = Number.isNaN(r.npv) ? "" : r.npv.toFixed(2);
        } else if (ind.key === "bcr") {
          value = Number.isNaN(r.bcr) ? "" : r.bcr.toFixed(4);
        } else if (ind.key === "roi") {
          value = Number.isNaN(r.roi) ? "" : r.roi.toFixed(2);
        } else if (ind.key === "rank") {
          value = r.rank;
        } else if (ind.key === "deltaNpv") {
          value = Number.isNaN(r.deltaNpv) ? "" : r.deltaNpv.toFixed(2);
        } else if (ind.key === "deltaPvCosts") {
          value = Number.isNaN(r.deltaPvCosts) ? "" : r.deltaPvCosts.toFixed(2);
        }
        row.push(value);
      }
      lines.push(row.map(csvEscape).join(","));
    }

    const blob = new Blob([lines.join("\n")], {
      type: "text/csv;charset=utf-8;"
    });
    downloadBlob(blob, "comparison_to_control.csv");
    showToast("Comparison-to-control results (CSV) downloaded.", "success");
  }

  function csvEscape(value) {
    if (value === null || value === undefined) return "";
    const s = String(value);
    if (s.includes(",") || s.includes('"') || s.includes("\n")) {
      return `"${s.replace(/"/g, '""')}"`;
    }
    return s;
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }, 0);
  }

  function exportWorkbook() {
    if (!window.XLSX) {
      showToast("Excel library not available.", "error");
      return;
    }
    if (!state.headers.length || !state.rows.length) {
      showToast("No dataset loaded for workbook export.", "error");
      return;
    }
    const wb = XLSX.utils.book_new();

    // Sheet 1: Cleaned dataset
    const dataSheetAoA = [state.headers.slice()];
    for (const row of state.rows) {
      dataSheetAoA.push(
        state.headers.map((h) =>
          row[h] === undefined || row[h] === null ? "" : row[h]
        )
      );
    }
    const wsData = XLSX.utils.aoa_to_sheet(dataSheetAoA);
    XLSX.utils.book_append_sheet(wb, wsData, "Dataset");

    // Sheet 2: Treatment summary
    const { treatments } = state.results;
    if (treatments && treatments.length) {
      const summaryAoA = [
        [
          "Treatment name",
          "Is control",
          "Average yield (t per ha)",
          "Average variable cost (per ha)",
          "Average capital cost (per ha)",
          "Total benefits over time (discounted)",
          "Total costs over time (discounted)",
          "Net profit over time",
          "Benefit per dollar spent",
          "Return on investment (percent)",
          "Rank",
          "Difference in net profit vs control",
          "Difference in total cost vs control"
        ]
      ];

      for (const t of treatments) {
        summaryAoA.push([
          t.name,
          t.isControl ? "TRUE" : "FALSE",
          Number.isNaN(t.avgYield) ? "" : t.avgYield,
          Number.isNaN(t.avgVarCost) ? "" : t.avgVarCost,
          Number.isNaN(t.avgCapCost) ? "" : t.avgCapCost,
          Number.isNaN(t.pvBenefits) ? "" : t.pvBenefits,
          Number.isNaN(t.pvTotalCosts) ? "" : t.pvTotalCosts,
          Number.isNaN(t.npv) ? "" : t.npv,
          Number.isNaN(t.bcr) ? "" : t.bcr,
          Number.isNaN(t.roi) ? "" : t.roi,
          t.rank,
          Number.isNaN(t.deltaNpv) ? "" : t.deltaNpv,
          Number.isNaN(t.deltaPvCosts) ? "" : t.deltaPvCosts
        ]);
      }

      const wsSummary = XLSX.utils.aoa_to_sheet(summaryAoA);
      XLSX.utils.book_append_sheet(wb, wsSummary, "Treatment summary");
    }

    // Sheet 3: Comparison to control
    const compAoA = [];
    const indicators = [
      { key: "pvBenefits", label: "Total benefits over time (discounted)" },
      { key: "pvTotalCosts", label: "Total costs over time (discounted)" },
      { key: "npv", label: "Net profit over time" },
      { key: "bcr", label: "Benefit per dollar spent" },
      { key: "roi", label: "Return on investment (percent)" },
      { key: "rank", label: "Overall ranking" },
      {
        key: "deltaNpv",
        label: "Difference in net profit compared with control"
      },
      {
        key: "deltaPvCosts",
        label: "Difference in total cost compared with control"
      }
    ];

    if (treatments && treatments.length) {
      const control = treatments.find((t) => t.isControl) || treatments[0];
      const nonControl = treatments.filter((t) => !t.isControl);
      const ordered = [control, ...nonControl];

      const headerRow = ["What is measured"];
      for (const t of ordered) {
        headerRow.push(t.isControl ? "Control (baseline)" : t.name);
      }
      compAoA.push(headerRow);

      for (const ind of indicators) {
        const row = [ind.label];
        for (const t of ordered) {
          const r = treatments.find((x) => x.name === t.name);
          let v = "";
          if (ind.key === "pvBenefits") v = r.pvBenefits;
          else if (ind.key === "pvTotalCosts") v = r.pvTotalCosts;
          else if (ind.key === "npv") v = r.npv;
          else if (ind.key === "bcr") v = r.bcr;
          else if (ind.key === "roi") v = r.roi;
          else if (ind.key === "rank") v = r.rank;
          else if (ind.key === "deltaNpv") v = r.deltaNpv;
          else if (ind.key === "deltaPvCosts") v = r.deltaPvCosts;
          row.push(v);
        }
        compAoA.push(row);
      }

      const wsComp = XLSX.utils.aoa_to_sheet(compAoA);
      XLSX.utils.book_append_sheet(wb, wsComp, "Comparison to control");
    }

    XLSX.writeFile(wb, "faba_beans_cba_results.xlsx");
    showToast("Excel workbook downloaded.", "success");
  }

  // =========================
  // 6) DATA LOAD & COMMIT
  // =========================

  async function loadDefaultDataset() {
    try {
      const response = await fetch("faba_beans_trial_clean_named.tsv", {
        cache: "no-cache"
      });
      if (!response.ok) {
        throw new Error("Default dataset could not be fetched.");
      }
      const text = await response.text();
      commitParsedData(text, "Default trial data loaded.");
    } catch (err) {
      console.error(err);
      showToast(
        "Could not load default dataset. You can still upload or paste data.",
        "error"
      );
    }
  }

  function commitParsedData(text, successMessage) {
    const parsed = parseDelimitedText(text);
    state.rawText = text;
    state.headers = parsed.headers;
    state.rows = parsed.rows;

    aggregateTreatments();
    computeCBA();
    renderAll();

    if (successMessage) {
      showToast(successMessage, "success");
    }
  }

  // =========================
  // 7) UI WIRING
  // =========================

  function onTabClick(event) {
    const button = event.currentTarget;
    const tabId = button.getAttribute("data-tab");
    if (!tabId) return;

    document.querySelectorAll(".tab-button").forEach((btn) => {
      btn.classList.toggle("active", btn === button);
    });
    document.querySelectorAll(".tab-panel").forEach((panel) => {
      panel.classList.toggle("active", panel.id === tabId);
    });
  }

  function onApplyScenario() {
    const priceInput = document.getElementById("pricePerTonne");
    const yearsInput = document.getElementById("years");
    const persInput = document.getElementById("persistenceYears");
    const discInput = document.getElementById("discountRate");
    const controlSelect = document.getElementById("controlChoice");

    if (priceInput) {
      state.params.pricePerTonne = parseNumber(priceInput.value) || 0;
    }
    if (yearsInput) {
      state.params.years = parseInt(yearsInput.value, 10) || 1;
    }
    if (persInput) {
      state.params.persistenceYears =
        parseInt(persInput.value, 10) || state.params.years;
    }
    if (discInput) {
      state.params.discountRate = parseNumber(discInput.value) || 0;
    }
    if (controlSelect && controlSelect.value) {
      state.controlName = controlSelect.value;
    }

    computeCBA();
    renderAll();
    showToast("Scenario settings updated.", "success");
  }

  function onFileInputChange(event) {
    const file = event.target.files && event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        commitParsedData(String(e.target.result), "Uploaded dataset loaded.");
      } catch (err) {
        console.error(err);
        showToast("Could not parse uploaded file. Check its format.", "error");
      }
    };
    reader.readAsText(file);
  }

  function onLoadPastedClick() {
    const ta = document.getElementById("pasteInput");
    if (!ta) return;
    const text = ta.value;
    if (!text || !text.trim()) {
      showToast("Paste data before loading.", "error");
      return;
    }
    try {
      commitParsedData(text, "Pasted dataset loaded.");
    } catch (err) {
      console.error(err);
      showToast("Could not parse pasted data. Check its format.", "error");
    }
  }

  function onReloadDefaultClick() {
    loadDefaultDataset();
  }

  function onLeaderboardFilterChange() {
    renderLeaderboard();
  }

  function onCopyAiBriefing() {
    const ta = document.getElementById("aiBriefing");
    if (!ta) return;
    ta.select();
    ta.setSelectionRange(0, ta.value.length);
    try {
      document.execCommand("copy");
      showToast("AI summary prompt copied.", "success");
    } catch (_) {
      showToast("Could not copy. Select and copy manually.", "error");
    }
  }

  function attachEventListeners() {
    document.querySelectorAll(".tab-button").forEach((btn) => {
      btn.addEventListener("click", onTabClick);
    });

    const applyBtn = document.getElementById("applyScenario");
    if (applyBtn) applyBtn.addEventListener("click", onApplyScenario);

    const fileInput = document.getElementById("fileInput");
    if (fileInput) fileInput.addEventListener("change", onFileInputChange);

    const btnLoadPasted = document.getElementById("btnLoadPasted");
    if (btnLoadPasted)
      btnLoadPasted.addEventListener("click", onLoadPastedClick);

    const btnReloadDefault = document.getElementById("btnReloadDefault");
    if (btnReloadDefault)
      btnReloadDefault.addEventListener("click", onReloadDefaultClick);

    const filterSelect = document.getElementById("leaderboardFilter");
    if (filterSelect)
      filterSelect.addEventListener("change", onLeaderboardFilterChange);

    const btnCleanData = document.getElementById("btnExportCleanData");
    if (btnCleanData)
      btnCleanData.addEventListener("click", exportCleanDatasetTSV);

    const btnSummary = document.getElementById("btnExportSummary");
    if (btnSummary)
      btnSummary.addEventListener("click", exportTreatmentSummaryCSV);

    const btnResults = document.getElementById("btnExportResults");
    if (btnResults)
      btnResults.addEventListener("click", exportComparisonCSV);

    const btnWorkbook = document.getElementById("btnExportWorkbook");
    if (btnWorkbook)
      btnWorkbook.addEventListener("click", exportWorkbook);

    const btnCopyAi = document.getElementById("btnCopyAiBriefing");
    if (btnCopyAi)
      btnCopyAi.addEventListener("click", onCopyAiBriefing);

    // Live scenario updates
    const priceInput = document.getElementById("pricePerTonne");
    const yearsInput = document.getElementById("years");
    const persInput = document.getElementById("persistenceYears");
    const discInput = document.getElementById("discountRate");
    const controlSelect = document.getElementById("controlChoice");

    if (priceInput)
      priceInput.addEventListener("change", onApplyScenario);
    if (yearsInput)
      yearsInput.addEventListener("change", onApplyScenario);
    if (persInput)
      persInput.addEventListener("change", onApplyScenario);
    if (discInput)
      discInput.addEventListener("change", onApplyScenario);
    if (controlSelect)
      controlSelect.addEventListener("change", onApplyScenario);
  }

  function renderAll() {
    renderOverview();
    renderDataSummaryAndChecks();
    renderControlChoice();
    computeCBA();
    renderLeaderboard();
    renderComparisonTable();
    renderCharts();
    buildAiBriefingPrompt();
  }

  document.addEventListener("DOMContentLoaded", () => {
    attachEventListeners();
    loadDefaultDataset();
  });
})();
