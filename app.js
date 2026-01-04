// app.js
// Farming CBA Decision Tool - Newcastle Business School
// World Bank style control centric cost benefit engine with:
// - Default faba beans trial loaded at start
// - Upload and paste import pipeline with basic validation
// - Replicate based control baselines where possible
// - Discounted net present value, benefit to cost ratio and return on investment
// - Comparison to control view with clear green and red cues
// - Sensitivity grid over discount rates and price multipliers
// - Scenario save and load with localStorage
// - Workbook and CSV exports
// - AI ready narrative and compact data block
// - Bottom right toast messages for key actions

(() => {
  "use strict";

  const SCENARIO_KEY = "farmingCbaScenarios_v1";

  const state = {
    rawText: "",
    header: [],
    rows: [],
    columns: {},
    treatments: [],
    controlName: null,
    config: {
      pricePerUnit: 600,
      horizonYears: 10,
      discountRate: 0.06,
      discountLow: 0.04,
      discountBase: 0.06,
      discountHigh: 0.08,
      priceLow: 0.8,
      priceBase: 1.0,
      priceHigh: 1.2
    },
    results: null,
    sensitivity: null,
    scenarios: []
  };

  // ======================
  // Helper: toast messages
  // ======================

  function showToast(message, type = "success") {
    const host = document.getElementById("toastHost");
    if (!host) return;

    const toast = document.createElement("div");
    toast.className = "toast " + (type === "error" ? "toast-error" : "toast-success");

    const span = document.createElement("span");
    span.className = "toast-message";
    span.textContent = message;

    const close = document.createElement("button");
    close.className = "toast-close";
    close.type = "button";
    close.textContent = "×";
    close.addEventListener("click", () => {
      toast.style.animation = "toast-out 0.14s ease-in forwards";
      setTimeout(() => toast.remove(), 140);
    });

    toast.appendChild(span);
    toast.appendChild(close);
    host.appendChild(toast);

    setTimeout(() => {
      if (!toast.isConnected) return;
      toast.style.animation = "toast-out 0.14s ease-in forwards";
      setTimeout(() => toast.remove(), 160);
    }, 3500);
  }

  // ======================
  // Helper: tabs
  // ======================

  function initTabs() {
    const tabs = document.querySelectorAll(".tab");
    const panels = document.querySelectorAll(".tab-panel");

    if (!tabs.length || !panels.length) return;

    const activate = (targetId) => {
      panels.forEach((panel) => {
        panel.classList.toggle("active", panel.id === targetId);
      });
      tabs.forEach((tab) => {
        tab.classList.toggle("active", tab.dataset.tabTarget === targetId);
      });
    };

    tabs.forEach((tab) => {
      tab.addEventListener("click", () => {
        const targetId = tab.dataset.tabTarget;
        if (!targetId) return;
        activate(targetId);
      });
    });

    // Ensure first tab is active on load
    const firstTab = tabs[0];
    if (firstTab) {
      activate(firstTab.dataset.tabTarget);
    }
  }

  // ======================
  // Helper: parsing
  // ======================

  function normaliseName(name) {
    return String(name || "")
      .trim()
      .toLowerCase()
      .replace(/\s+/g, "_");
  }

  function detectDelimiter(line) {
    if (!line) return "\t";
    const tabCount = (line.match(/\t/g) || []).length;
    const commaCount = (line.match(/,/g) || []).length;
    if (tabCount >= commaCount) return "\t";
    return ",";
  }

  function parseTabular(text) {
    const cleaned = String(text || "").replace(/\r\n/g, "\n").replace(/\r/g, "\n").trim();
    if (!cleaned) {
      return { header: [], rows: [], delimiter: "\t" };
    }

    const lines = cleaned.split("\n").filter((l) => l.trim().length > 0);

    let headerIndex = 0;
    // Skip simple dictionary style lines if any
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();
      if (!line) continue;
      if (/\breplicate\b/i.test(line) && /\btreatment\b/i.test(line)) {
        headerIndex = i;
        break;
      }
    }

    const headerLine = lines[headerIndex];
    const delimiter = detectDelimiter(headerLine);
    const headerRaw = headerLine.split(delimiter);
    const header = headerRaw.map((h) => h.trim());

    const rows = [];
    for (let i = headerIndex + 1; i < lines.length; i++) {
      const line = lines[i];
      if (!line || !line.trim()) continue;
      const parts = line.split(delimiter);
      if (parts.every((p) => p.trim() === "")) continue;
      const row = {};
      header.forEach((h, idx) => {
        row[h] = parts[idx] !== undefined ? parts[idx].trim() : "";
      });
      rows.push(row);
    }

    return { header, rows, delimiter };
  }

  function buildColumns(header) {
    const cols = {};
    header.forEach((h, idx) => {
      cols[normaliseName(h)] = { index: idx, name: h };
    });
    return cols;
  }

  function getColumnIndex(key) {
    const wanted = Array.isArray(key) ? key : [key];
    for (const name of wanted) {
      const norm = normaliseName(name);
      if (Object.prototype.hasOwnProperty.call(state.columns, norm)) {
        return state.columns[norm].index;
      }
    }
    return -1;
  }

  // ======================
  // Data import and commit
  // ======================

  function commitDataset(text, sourceLabel) {
    const parsed = parseTabular(text);
    state.rawText = text;
    state.header = parsed.header;
    state.rows = parsed.rows;
    state.columns = buildColumns(parsed.header);

    detectTreatmentsAndControl();
    const checks = buildDataChecks();
    renderDataChecks(checks);
    renderVariableSummary();
    runAllCalculations();

    const pill = document.getElementById("datasetStatusPill");
    if (pill) {
      const valueSpan = pill.querySelector(".value");
      if (valueSpan) {
        valueSpan.textContent = sourceLabel + " loaded";
      }
    }

    updateImportSummary(sourceLabel, checks);
    showToast("Data loaded from " + sourceLabel);
  }

  // ======================
  // Treatment detection
  // ======================

  function detectTreatmentsAndControl() {
    const h = state.header;
    const rows = state.rows;

    if (!h.length || !rows.length) {
      state.treatments = [];
      state.controlName = null;
      return;
    }

    const treatmentIdx = getColumnIndex(["treatment", "treatment_name", "treat"]);
    if (treatmentIdx === -1) {
      // Build pseudo treatments
      state.treatments = [];
      state.controlName = null;
      return;
    }

    const controlIdx = getColumnIndex(["is_control", "control", "control_flag"]);
    const replicateIdx = getColumnIndex(["replicate", "rep"]);

    const treatmentSet = new Set();
    const controlSet = new Set();

    rows.forEach((row) => {
      const values = Object.values(row);
      const treat = values[treatmentIdx] || "";
      if (!treat) return;
      treatmentSet.add(treat);
      if (controlIdx !== -1) {
        const ctrlRaw = values[controlIdx];
        const flag = String(ctrlRaw || "")
          .trim()
          .toLowerCase();
        if (flag === "1" || flag === "yes" || flag === "true" || flag === "control") {
          controlSet.add(treat);
        }
      }
    });

    const treatments = Array.from(treatmentSet);
    let controlName = null;

    if (controlSet.size === 1) {
      controlName = Array.from(controlSet)[0];
    } else if (controlSet.size > 1) {
      // Pick the most frequent control
      const counts = {};
      state.rows.forEach((row) => {
        const values = Object.values(row);
        const treat = values[treatmentIdx] || "";
        if (!treat) return;
        const ctrlRaw = controlIdx !== -1 ? values[controlIdx] : "";
        const flag = String(ctrlRaw || "")
          .trim()
          .toLowerCase();
        if (flag === "1" || flag === "yes" || flag === "true" || flag === "control") {
          counts[treat] = (counts[treat] || 0) + 1;
        }
      });
      let best = null;
      let bestCount = -1;
      Object.entries(counts).forEach(([t, c]) => {
        if (c > bestCount) {
          best = t;
          bestCount = c;
        }
      });
      controlName = best;
    } else if (treatments.length) {
      // Fallback: use first treatment as control
      controlName = treatments[0];
    }

    state.treatments = treatments;
    state.controlName = controlName;
  }

  // ======================
  // Data checks and summary
  // ======================

  function buildDataChecks() {
    const checks = [];
    const h = state.header;
    const rows = state.rows;

    if (!h.length || !rows.length) {
      checks.push({
        type: "error",
        label: "No data rows found",
        count: 0,
        details: "Upload or paste a dataset with at least one header row and one data row."
      });
      return checks;
    }

    const treatmentIdx = getColumnIndex(["treatment", "treatment_name", "treat"]);
    const replicateIdx = getColumnIndex(["replicate", "rep"]);
    const yieldIdx = getColumnIndex(["yield", "yield_kg", "yield_t_ha"]);
    const varCostIdx = getColumnIndex(["variable_cost", "var_cost", "cost_variable"]);
    const capCostIdx = getColumnIndex(["capital_cost", "cap_cost", "fixed_cost"]);
    const controlIdx = getColumnIndex(["is_control", "control", "control_flag"]);

    const missingCore = [];
    if (treatmentIdx === -1) missingCore.push("treatment name");
    if (yieldIdx === -1) missingCore.push("yield");
    if (varCostIdx === -1 && capCostIdx === -1) missingCore.push("any cost field");

    if (missingCore.length) {
      checks.push({
        type: "error",
        label: "Key fields missing",
        count: missingCore.length,
        details:
          "The following fields were not detected: " +
          missingCore.join(", ") +
          ". The tool can still show basic tables but the economic engine will be limited."
      });
    }

    if (replicateIdx === -1) {
      checks.push({
        type: "warning",
        label: "Replicate field missing",
        count: 1,
        details:
          "No clear replicate field was found. The tool will treat the dataset as if there was one large block, and will not use replicate based control baselines."
      });
    }

    if (!state.controlName) {
      checks.push({
        type: "warning",
        label: "Control not clearly identified",
        count: 1,
        details:
          "There was no unique control flag. The tool has selected one treatment as the control based on the data, but you may want to check that this matches your trial design."
      });
    }

    // Simple missing value count on core numeric fields
    let missingNumeric = 0;
    if (yieldIdx !== -1 || varCostIdx !== -1 || capCostIdx !== -1) {
      rows.forEach((row) => {
        const values = Object.values(row);
        if (yieldIdx !== -1) {
          const v = values[yieldIdx];
          if (v === "" || v === "?" || isNaN(parseFloat(v))) missingNumeric++;
        }
        if (varCostIdx !== -1) {
          const v = values[varCostIdx];
          if (v === "" || v === "?" || isNaN(parseFloat(v))) missingNumeric++;
        }
        if (capCostIdx !== -1) {
          const v = values[capCostIdx];
          if (v === "" || v === "?" || isNaN(parseFloat(v))) missingNumeric++;
        }
      });
    }

    if (missingNumeric > 0) {
      checks.push({
        type: "warning",
        label: "Missing or non numeric values",
        count: missingNumeric,
        details:
          "Some yield or cost values are missing or non numeric. These rows are ignored in the calculations but kept in the cleaned export file."
      });
    }

    // Basic row count
    checks.push({
      type: "info",
      label: "Row and treatment count",
      count: rows.length,
      details:
        "The dataset has " +
        rows.length +
        " rows and " +
        state.treatments.length +
        " treatments. The current control is: " +
        (state.controlName || "not set") +
        "."
    });

    return checks;
  }

  function renderDataChecks(checks) {
    const container = document.getElementById("dataChecks");
    if (!container) return;
    container.innerHTML = "";

    if (!checks || !checks.length) {
      return;
    }

    checks.forEach((check) => {
      const div = document.createElement("div");
      div.className = "check-row " + (check.type === "error" ? "error" : check.type === "warning" ? "warning" : "");
      const label = document.createElement("span");
      label.className = "check-label";
      label.textContent = check.label;

      const tag = document.createElement("span");
      tag.className = "check-tag " + (check.type === "error" ? "error" : check.type === "warning" ? "warning" : "");
      tag.textContent = check.type === "error" ? "Check" : check.type === "warning" ? "Notice" : "Info";

      const count = document.createElement("span");
      count.className = "check-count";
      count.textContent = "Count: " + check.count;

      const details = document.createElement("span");
      details.className = "check-count";
      details.textContent = check.details;

      div.appendChild(label);
      div.appendChild(tag);
      div.appendChild(count);
      div.appendChild(details);
      container.appendChild(div);
    });
  }

  function renderVariableSummary() {
    const container = document.getElementById("variableSummary");
    if (!container) return;

    if (!state.header.length) {
      container.innerHTML = "<p class='small muted'>No header row detected.</p>";
      return;
    }

    const keyMap = {
      replicate: ["replicate", "rep"],
      treatment: ["treatment", "treatment_name", "treat"],
      control_flag: ["is_control", "control_flag", "control"],
      yield: ["yield", "yield_kg", "yield_t_ha"],
      variable_cost: ["variable_cost", "var_cost", "cost_variable"],
      capital_cost: ["capital_cost", "cap_cost", "fixed_cost"]
    };

    let html = "<table class='summary-table'><thead><tr><th>Purpose</th><th>Column used</th></tr></thead><tbody>";
    Object.entries(keyMap).forEach(([purpose, aliases]) => {
      const idx = getColumnIndex(aliases);
      const used = idx === -1 ? "<span class='muted'>Not found</span>" : state.header[idx];
      const label =
        purpose === "replicate"
          ? "Replicate"
          : purpose === "treatment"
          ? "Treatment name"
          : purpose === "control_flag"
          ? "Control flag"
          : purpose === "yield"
          ? "Yield"
          : purpose === "variable_cost"
          ? "Variable cost"
          : "Capital cost";
      html += "<tr><td>" + label + "</td><td>" + used + "</td></tr>";
    });
    html += "</tbody></table>";

    container.innerHTML = html;
  }

  function updateImportSummary(sourceLabel, checks) {
    const el = document.getElementById("importSummary");
    if (!el) return;
    const errors = checks.filter((c) => c.type === "error").length;
    const warnings = checks.filter((c) => c.type === "warning").length;

    let msg =
      "Data loaded from " +
      sourceLabel +
      ". There are " +
      errors +
      " key issues and " +
      warnings +
      " warnings. Scroll down to view details.";
    el.textContent = msg;
  }

  // ======================
  // Economic engine
  // ======================

  function parseNumberSafe(value) {
    if (value === null || value === undefined) return null;
    const s = String(value).trim();
    if (s === "" || s === "?") return null;
    const n = parseFloat(s);
    return isNaN(n) ? null : n;
  }

  function computeTreatmentStatistics(configOverride) {
    const rows = state.rows;
    const header = state.header;
    if (!rows.length || !header.length || !state.treatments.length) {
      return [];
    }

    const cfg = configOverride || state.config;

    const treatmentIdx = getColumnIndex(["treatment", "treatment_name", "treat"]);
    const replicateIdx = getColumnIndex(["replicate", "rep"]);
    const yieldIdx = getColumnIndex(["yield", "yield_kg", "yield_t_ha"]);
    const varCostIdx = getColumnIndex(["variable_cost", "var_cost", "cost_variable"]);
    const capCostIdx = getColumnIndex(["capital_cost", "cap_cost", "fixed_cost"]);

    if (treatmentIdx === -1 || yieldIdx === -1) {
      return [];
    }

    // Map replicate -> control yield
    const controlName = state.controlName;
    const repControlYield = {};
    if (controlName && replicateIdx !== -1) {
      rows.forEach((row) => {
        const values = Object.values(row);
        const treat = values[treatmentIdx];
        const rep = values[replicateIdx];
        if (!treat || !rep) return;
        if (treat !== controlName) return;
        const y = parseNumberSafe(values[yieldIdx]);
        if (y === null) return;
        repControlYield[rep] = y;
      });
    }

    const stats = {};
    state.treatments.forEach((t) => {
      stats[t] = {
        treatment: t,
        isControl: t === controlName,
        plots: 0,
        replicates: new Set(),
        yields: [],
        yieldDelta: [],
        varCosts: [],
        capCosts: []
      };
    });

    rows.forEach((row) => {
      const values = Object.values(row);
      const treat = values[treatmentIdx];
      if (!treat || !stats[treat]) return;
      const st = stats[treat];
      const rep = replicateIdx !== -1 ? values[replicateIdx] : "all";
      const y = parseNumberSafe(values[yieldIdx]);
      const vc = varCostIdx !== -1 ? parseNumberSafe(values[varCostIdx]) : null;
      const cc = capCostIdx !== -1 ? parseNumberSafe(values[capCostIdx]) : null;

      st.plots += 1;
      st.replicates.add(rep);
      if (y !== null) st.yields.push(y);
      if (vc !== null) st.varCosts.push(vc);
      if (cc !== null) st.capCosts.push(cc);

      if (!st.isControl && repControlYield[rep] !== undefined && y !== null) {
        st.yieldDelta.push(y - repControlYield[rep]);
      }
    });

    const treatmentsStats = [];
    const price = cfg.pricePerUnit;
    const T = cfg.horizonYears;
    const r = cfg.discountRate;

    let pvFactor = T;
    if (r > 0) {
      pvFactor = (1 - Math.pow(1 + r, -T)) / r;
    }

    Object.values(stats).forEach((st) => {
      const avgYield =
        st.yields.length > 0
          ? st.yields.reduce((a, b) => a + b, 0) / st.yields.length
          : null;
      const avgVarCost =
        st.varCosts.length > 0
          ? st.varCosts.reduce((a, b) => a + b, 0) / st.varCosts.length
          : 0;
      const avgCapCost =
        st.capCosts.length > 0
          ? st.capCosts.reduce((a, b) => a + b, 0) / st.capCosts.length
          : 0;

      const avgDeltaYield =
        st.yieldDelta.length > 0
          ? st.yieldDelta.reduce((a, b) => a + b, 0) / st.yieldDelta.length
          : 0;

      const annualBenefit = avgDeltaYield * price;
      const pvBenefits = annualBenefit * pvFactor;
      const pvVarCost = avgVarCost * pvFactor;
      const pvCosts = pvVarCost + avgCapCost;

      const npv = pvBenefits - pvCosts;
      const bcr = pvCosts > 0 ? pvBenefits / pvCosts : null;
      const roi = pvCosts > 0 ? npv / pvCosts : null;

      treatmentsStats.push({
        treatment: st.treatment,
        isControl: st.isControl,
        plots: st.plots,
        replicateCount: st.replicates.size,
        avgYield,
        avgDeltaYield,
        avgVarCost,
        avgCapCost,
        pvBenefits,
        pvCosts,
        npv,
        bcr,
        roi
      });
    });

    // Rank by NPV descending
    treatmentsStats.sort((a, b) => (b.npv || 0) - (a.npv || 0));
    treatmentsStats.forEach((st, idx) => {
      st.rank = idx + 1;
    });

    return treatmentsStats;
  }

  function formatCurrency(value) {
    if (value === null || value === undefined || isNaN(value)) return "";
    const abs = Math.abs(value);
    const decimals = abs >= 1000 ? 0 : 2;
    return value.toLocaleString(undefined, {
      minimumFractionDigits: decimals,
      maximumFractionDigits: decimals
    });
  }

  function formatRatio(value) {
    if (value === null || value === undefined || isNaN(value)) return "";
    return value.toLocaleString(undefined, {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    });
  }

  function runAllCalculations() {
    const treatmentsStats = computeTreatmentStatistics();
    state.results = { treatments: treatmentsStats };
    renderTreatmentsView(treatmentsStats);
    renderResultsLeaderboard(treatmentsStats);
    renderComparisonTable(treatmentsStats);
    renderNarrative(treatmentsStats);
    computeAndRenderSensitivity();
    fillAiBriefing(treatmentsStats);
    fillStructuredResults(treatmentsStats);
  }

  // ======================
  // Treatments view
  // ======================

  function renderTreatmentsView(treatmentsStats) {
    const container = document.getElementById("treatmentsList");
    if (!container) return;

    if (!treatmentsStats || !treatmentsStats.length) {
      container.innerHTML =
        "<p class='small muted'>No treatment level summaries are available yet. Check the Data and checks tab to make sure the key fields are present.</p>";
      return;
    }

    container.innerHTML = "";
    treatmentsStats.forEach((st) => {
      const card = document.createElement("div");
      card.className = "treatment-card";

      const left = document.createElement("div");
      const right = document.createElement("div");

      const header = document.createElement("div");
      header.className = "treatment-card-header";

      const nameSpan = document.createElement("span");
      nameSpan.className = "treatment-name";
      nameSpan.textContent = st.treatment;

      const tagSpan = document.createElement("span");
      tagSpan.className = "treatment-tag" + (st.isControl ? " control" : "");
      tagSpan.textContent = st.isControl ? "Control" : "Treatment";

      header.appendChild(nameSpan);
      header.appendChild(tagSpan);

      const meta = document.createElement("div");
      meta.className = "treatment-meta";
      meta.textContent =
        "Plots: " +
        st.plots +
        " · Replicates: " +
        st.replicateCount +
        (st.isControl ? " · Used as baseline" : "");

      left.appendChild(header);
      left.appendChild(meta);

      const yieldTitle = document.createElement("div");
      yieldTitle.className = "treatment-section-title";
      yieldTitle.textContent = "Average yield and yield change";

      const costTitle = document.createElement("div");
      costTitle.className = "treatment-section-title";
      costTitle.style.marginTop = "8px";
      costTitle.textContent = "Average costs per plot";

      const table1 = document.createElement("table");
      table1.className = "treatment-summary-table";
      table1.innerHTML =
        "<tbody>" +
        "<tr><th>Average yield</th><td>" +
        (st.avgYield !== null && !isNaN(st.avgYield)
          ? st.avgYield.toFixed(3)
          : "") +
        "</td></tr>" +
        "<tr><th>Average change vs control</th><td>" +
        (st.avgDeltaYield !== null && !isNaN(st.avgDeltaYield)
          ? st.avgDeltaYield.toFixed(3)
          : st.isControl
          ? "0.000"
          : "") +
        "</td></tr>" +
        "</tbody>";

      const table2 = document.createElement("table");
      table2.className = "treatment-summary-table";
      table2.innerHTML =
        "<tbody>" +
        "<tr><th>Average variable cost</th><td>" +
        formatCurrency(st.avgVarCost) +
        "</td></tr>" +
        "<tr><th>Average capital cost</th><td>" +
        formatCurrency(st.avgCapCost) +
        "</td></tr>" +
        "</tbody>";

      right.appendChild(yieldTitle);
      right.appendChild(table1);
      right.appendChild(costTitle);
      right.appendChild(table2);

      card.appendChild(left);
      card.appendChild(right);
      container.appendChild(card);
    });
  }

  // ======================
  // Results leaderboard
  // ======================

  function renderResultsLeaderboard(treatmentsStats, filterType) {
    const container = document.getElementById("resultsLeaderboard");
    if (!container) return;

    if (!treatmentsStats || !treatmentsStats.length) {
      container.innerHTML =
        "<p class='small muted'>No economic results yet. Check the Configuration and scenarios tab and make sure key fields are present in the data.</p>";
      return;
    }

    let list = [...treatmentsStats];

    const controlName = state.controlName;
    const control = list.find((x) => x.treatment === controlName);

    if (filterType === "topNPV") {
      list.sort((a, b) => (b.npv || 0) - (a.npv || 0));
      list = list.slice(0, 5);
    } else if (filterType === "topBCR") {
      list.sort((a, b) => (b.bcr || 0) - (a.bcr || 0));
      list = list.slice(0, 5);
    } else if (filterType === "onlyBetter" && control) {
      list = list.filter((t) => t.npv > control.npv);
      if (!list.length) {
        container.innerHTML =
          "<p class='small muted'>No treatment has a higher net present value than the control under the current settings.</p>";
        return;
      }
    }

    const headers =
      "<thead><tr>" +
      "<th>Rank</th>" +
      "<th>Treatment</th>" +
      "<th>NPV</th>" +
      "<th>BCR</th>" +
      "<th>ROI</th>" +
      "<th>Change in NPV vs control</th>" +
      "</tr></thead>";

    let rowsHtml = "<tbody>";
    list.forEach((st) => {
      const isControl = st.treatment === controlName;
      const rankChipClass = isControl ? "" : st.rank <= 3 ? "rank-chip top" : "rank-chip";
      const rankChipText = isControl ? "-" : st.rank;
      const controlNpv = control ? control.npv : 0;
      const deltaNpv = st.npv - controlNpv;

      const deltaClass =
        !isControl && deltaNpv > 0
          ? "cell-positive"
          : !isControl && deltaNpv < 0
          ? "cell-negative"
          : "";

      rowsHtml +=
        "<tr class='" +
        (isControl ? "highlight" : "") +
        "'>" +
        "<td><span class='" +
        rankChipClass +
        "'>" +
        rankChipText +
        "</span></td>" +
        "<td>" +
        st.treatment +
        (isControl ? " (control)" : "") +
        "</td>" +
        "<td>" +
        formatCurrency(st.npv) +
        "</td>" +
        "<td>" +
        (st.bcr !== null ? formatRatio(st.bcr) : "") +
        "</td>" +
        "<td>" +
        (st.roi !== null ? formatRatio(st.roi) : "") +
        "</td>" +
        "<td class='" +
        deltaClass +
        "'>" +
        (isControl ? "0" : formatCurrency(deltaNpv)) +
        "</td>" +
        "</tr>";
    });
    rowsHtml += "</tbody>";

    container.innerHTML = "<table>" + headers + rowsHtml + "</table>";
  }

  // ======================
  // Comparison to control table
  // ======================

  function renderComparisonTable(treatmentsStats) {
    const container = document.getElementById("comparisonToControl");
    if (!container) return;

    if (!treatmentsStats || !treatmentsStats.length) {
      container.innerHTML =
        "<p class='small muted' style='padding:10px;'>No economic results yet. Once data and settings are in place, a full comparison table will appear here.</p>";
      return;
    }

    const controlName = state.controlName;
    const control = treatmentsStats.find((t) => t.treatment === controlName) || treatmentsStats[0];

    const indicators = [
      { key: "pvBenefits", label: "Present value of benefits" },
      { key: "pvCosts", label: "Present value of costs" },
      { key: "npv", label: "Net present value" },
      { key: "bcr", label: "Benefit to cost ratio", isRatio: true },
      { key: "roi", label: "Return on investment", isRatio: true }
    ];

    const headers = [];
    headers.push("<th>Indicator</th>");
    headers.push("<th class='header-control'>Control (baseline)</th>");
    treatmentsStats.forEach((st) => {
      if (st.treatment === control.treatment) return;
      headers.push("<th class='header-treatment'>" + st.treatment + "</th>");
      headers.push("<th class='header-treatment'>Change in NPV vs control</th>");
      headers.push("<th class='header-treatment'>Change in PV costs vs control</th>");
    });

    let html = "<table><thead><tr>" + headers.join("") + "</tr></thead><tbody>";

    indicators.forEach((ind) => {
      let row = "<tr>";
      row += "<th>" + ind.label + "</th>";

      const controlVal = control[ind.key];
      const controlStr = ind.isRatio ? formatRatio(controlVal) : formatCurrency(controlVal);
      row += "<td class='cell-strong'>" + controlStr + "</td>";

      treatmentsStats.forEach((st) => {
        if (st.treatment === control.treatment) return;
        const val = st[ind.key];
        const valStr = ind.isRatio ? formatRatio(val) : formatCurrency(val);

        if (ind.key === "npv" || ind.key === "pvBenefits" || ind.key === "pvCosts") {
          const deltaNpv = st.npv - control.npv;
          const deltaCosts = st.pvCosts - control.pvCosts;
          const deltaClassNpv =
            deltaNpv > 0 ? "cell-positive cell-strong" : deltaNpv < 0 ? "cell-negative cell-strong" : "";
          const deltaClassCost =
            deltaCosts < 0 ? "cell-positive cell-strong" : deltaCosts > 0 ? "cell-negative cell-strong" : "";
          if (ind.key === "npv") {
            row += "<td>" + valStr + "</td>";
            row += "<td class='" + deltaClassNpv + "'>" + formatCurrency(deltaNpv) + "</td>";
            row += "<td class='" + deltaClassCost + "'>" + formatCurrency(deltaCosts) + "</td>";
          } else if (ind.key === "pvBenefits") {
            row += "<td>" + valStr + "</td>";
            row += "<td class='" + deltaClassNpv + "'>" + formatCurrency(deltaNpv) + "</td>";
            row += "<td class='" + deltaClassCost + "'>" + formatCurrency(deltaCosts) + "</td>";
          } else if (ind.key === "pvCosts") {
            row += "<td>" + valStr + "</td>";
            row += "<td class='" + deltaClassNpv + "'>" + formatCurrency(deltaNpv) + "</td>";
            row += "<td class='" + deltaClassCost + "'>" + formatCurrency(deltaCosts) + "</td>";
          }
        } else {
          // Ratios and ROI: only show main value and leave delta cells blank
          row += "<td>" + valStr + "</td>";
          row += "<td></td>";
          row += "<td></td>";
        }
      });

      row += "</tr>";
      html += row;
    });

    html += "</tbody></table>";
    container.innerHTML = html;
  }

  // ======================
  // Narrative
  // ======================

  function renderNarrative(treatmentsStats) {
    const el = document.getElementById("resultsNarrative");
    if (!el) return;

    if (!treatmentsStats || !treatmentsStats.length) {
      el.textContent =
        "Once data and configuration are in place, this section will describe the main findings in plain language.";
      return;
    }

    const controlName = state.controlName;
    const control = treatmentsStats.find((t) => t.treatment === controlName) || treatmentsStats[0];
    const others = treatmentsStats.filter((t) => t.treatment !== control.treatment);

    if (!others.length) {
      el.textContent =
        "Only one treatment is present in the dataset. Add at least one more treatment for the comparison to become meaningful.";
      return;
    }

    const bestByNpv = [...others].sort((a, b) => (b.npv || 0) - (a.npv || 0))[0];
    const bestByBcr = [...others].sort((a, b) => (b.bcr || 0) - (a.bcr || 0))[0];

    const deltaNpv = bestByNpv.npv - control.npv;
    const deltaCost = bestByNpv.pvCosts - control.pvCosts;
    const directionNpv = deltaNpv > 0 ? "higher" : deltaNpv < 0 ? "lower" : "the same as";
    const directionCost = deltaCost > 0 ? "higher" : deltaCost < 0 ? "lower" : "the same as";

    const cfg = state.config;

    const text =
      "Under the current settings, with a price per unit of " +
      formatCurrency(cfg.pricePerUnit) +
      " and a " +
      (cfg.discountRate * 100).toFixed(1) +
      " percent discount rate over " +
      cfg.horizonYears +
      " years, the control is " +
      control.treatment +
      ". The treatment with the highest net present value is " +
      bestByNpv.treatment +
      ". Its net present value is " +
      formatCurrency(bestByNpv.npv) +
      ", which is " +
      directionNpv +
      " the control by about " +
      formatCurrency(Math.abs(deltaNpv)) +
      ". The present value of costs for this treatment is " +
      directionCost +
      " the control by around " +
      formatCurrency(Math.abs(deltaCost)) +
      ". This means that it wins mainly through " +
      (deltaNpv > 0 && deltaCost <= 0
        ? "a mix of higher benefits and lower costs"
        : deltaNpv > 0 && deltaCost > 0
        ? "higher benefits that more than cover higher costs"
        : "a different combination of costs and benefits") +
      ".\n\n" +
      "Looking at cost effectiveness, the treatment with the highest benefit to cost ratio is " +
      bestByBcr.treatment +
      " with a ratio of about " +
      formatRatio(bestByBcr.bcr) +
      ". This means that every unit of cost spent on this treatment returns that many units of benefit in present value terms. " +
      "Treatments with a ratio above one return more benefits than costs. " +
      "Treatments with a ratio below one return less benefits than costs under the current assumptions.";

    el.textContent = text;
  }

  // ======================
  // Sensitivity
  // ======================

  function computeAndRenderSensitivity() {
    const el = document.getElementById("sensitivitySummary");
    if (!el) return;

    const cfg = state.config;

    const discounts = [cfg.discountLow, cfg.discountBase, cfg.discountHigh];
    const prices = [cfg.priceLow, cfg.priceBase, cfg.priceHigh];

    const baseStats = computeTreatmentStatistics();
    if (!baseStats.length) {
      el.textContent =
        "Sensitivity results will appear here once the economic engine has been able to calculate net present values.";
      return;
    }

    const controlName = state.controlName;
    const controlBase =
      baseStats.find((t) => t.treatment === controlName) || baseStats[0];

    const bestBase = [...baseStats]
      .filter((t) => t.treatment !== controlBase.treatment)
      .sort((a, b) => (b.npv || 0) - (a.npv || 0))[0];

    if (!bestBase) {
      el.textContent =
        "Only one treatment is present in the dataset. Sensitivity analysis is not very informative in this case.";
      return;
    }

    const records = [];

    discounts.forEach((dr) => {
      prices.forEach((pm) => {
        const override = {
          ...cfg,
          discountRate: dr,
          pricePerUnit: cfg.pricePerUnit * pm
        };
        const stats = computeTreatmentStatistics(override);
        if (!stats.length) return;
        const controlAlt =
          stats.find((t) => t.treatment === controlBase.treatment) || stats[0];
        const bestAlt = [...stats]
          .filter((t) => t.treatment !== controlAlt.treatment)
          .sort((a, b) => (b.npv || 0) - (a.npv || 0))[0];
        if (!bestAlt) return;

        records.push({
          discount: dr,
          priceMultiplier: pm,
          bestTreatment: bestAlt.treatment,
          bestNpv: bestAlt.npv,
          controlNpv: controlAlt.npv
        });
      });
    });

    state.sensitivity = { records };

    let text =
      "The table below looks at combinations of discount rates and price multipliers. " +
      "For each combination it shows the treatment with the highest net present value and the margin relative to the control.\n\n";

    records.forEach((rec) => {
      const ratePct = (rec.discount * 100).toFixed(1);
      const pricePct = (rec.priceMultiplier * 100).toFixed(0);
      const delta = rec.bestNpv - rec.controlNpv;
      text +=
        "Discount rate " +
        ratePct +
        " percent and price at " +
        pricePct +
        " percent of the base level: " +
        rec.bestTreatment +
        " remains in first place with a net present value of " +
        formatCurrency(rec.bestNpv) +
        ". This is ahead of the control by about " +
        formatCurrency(delta) +
        ".\n";
    });

    el.textContent = text;
  }

  // ======================
  // AI briefing helpers
  // ======================

  function fillAiBriefing(treatmentsStats) {
    const el = document.getElementById("aiBriefingText");
    if (!el) return;

    if (!treatmentsStats || !treatmentsStats.length) {
      el.value =
        "The economic engine has not yet been able to calculate net present values. Once data and configuration are in place, this box will contain a ready to use description for an assistant.";
      return;
    }

    const cfg = state.config;
    const controlName = state.controlName;
    const control =
      treatmentsStats.find((t) => t.treatment === controlName) || treatmentsStats[0];

    const ordered = [...treatmentsStats].sort((a, b) => (b.npv || 0) - (a.npv || 0));
    const topTwo = ordered.slice(0, 2);

    let txt = "";
    txt +=
      "You are asked to prepare a clear narrative summary of a farm trial and its cost benefit results. " +
      "The trial compares one control treatment with several alternative treatments. The tool has already calculated net present values, benefit to cost ratios and returns on investment.\n\n";

    txt +=
      "Key economic settings are as follows. The analysis horizon is " +
      cfg.horizonYears +
      " years. The main discount rate is " +
      (cfg.discountRate * 100).toFixed(1) +
      " percent per year. The price per unit of output is " +
      formatCurrency(cfg.pricePerUnit) +
      " in local currency. Capital costs are treated as year zero costs and are not discounted. Variable costs and benefits are treated as annual flows and are discounted over the analysis horizon.\n\n";

    txt +=
      "The control treatment is called " +
      control.treatment +
      ". Under the main economic settings its net present value is " +
      formatCurrency(control.npv) +
      " in local currency. This provides a baseline for comparison.\n\n";

    txt +=
      "The treatments with the highest net present values are listed below with their net present values and benefit to cost ratios. " +
      "The first item is the control for reference.\n\n";

    const lines = [];
    treatmentsStats.forEach((st) => {
      lines.push(
        "Treatment name: " +
          st.treatment +
          ". Net present value: " +
          formatCurrency(st.npv) +
          ". Benefit to cost ratio: " +
          (st.bcr !== null ? formatRatio(st.bcr) : "not defined") +
          ". Return on investment: " +
          (st.roi !== null ? formatRatio(st.roi) : "not defined") +
          "."
      );
    });

    txt += lines.join("\n") + "\n\n";

    if (topTwo.length >= 2) {
      const first = topTwo[0];
      const second = topTwo[1];
      txt +=
        "In your narrative, highlight why " +
        first.treatment +
        " is ranked first based on net present value and how much it gains over the control. " +
        "Also explain how " +
        second.treatment +
        " compares with the control and with " +
        first.treatment +
        ".\n\n";
    }

    txt +=
      "Please write in plain language that is suitable for farmers and local advisers. " +
      "Explain what net present value, the benefit to cost ratio and return on investment mean in intuitive terms. " +
      "Make it clear that results depend on the price per unit and the discount rate. " +
      "Avoid technical jargon wherever possible.";

    el.value = txt;
  }

  function fillStructuredResults(treatmentsStats) {
    const el = document.getElementById("resultsJson");
    if (!el) return;

    if (!treatmentsStats || !treatmentsStats.length) {
      el.value =
        '{"note":"No results yet. Once net present values are available this block will contain a compact summary."}';
      return;
    }

    const cfg = state.config;
    const data = {
      meta: {
        description: "Farming CBA decision aid summary prepared for an assistant",
        horizonYears: cfg.horizonYears,
        discountRate: cfg.discountRate,
        pricePerUnit: cfg.pricePerUnit
      },
      treatments: treatmentsStats.map((st) => ({
        name: st.treatment,
        isControl: st.isControl,
        plots: st.plots,
        replicateCount: st.replicateCount,
        avgYield: st.avgYield,
        avgDeltaYield: st.avgDeltaYield,
        avgVarCost: st.avgVarCost,
        avgCapCost: st.avgCapCost,
        pvBenefits: st.pvBenefits,
        pvCosts: st.pvCosts,
        npv: st.npv,
        bcr: st.bcr,
        roi: st.roi,
        rank: st.rank
      }))
    };

    el.value = JSON.stringify(data, null, 2);
  }

  // ======================
  // Exports
  // ======================

  function downloadFile(filename, content, mime) {
    const blob = new Blob([content], { type: mime || "text/plain;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  function exportCleanTsv() {
    if (!state.header.length || !state.rows.length) {
      showToast("No data to export.", "error");
      return;
    }

    const delimiter = "\t";
    const headerLine = state.header.join(delimiter);
    const lines = [headerLine];

    state.rows.forEach((row) => {
      const parts = state.header.map((h) => {
        const v = row[h];
        return v === undefined || v === null ? "" : String(v).replace(/\r?\n/g, " ");
      });
      lines.push(parts.join(delimiter));
    });

    const content = lines.join("\n");
    downloadFile("farming_cba_cleaned.tsv", content, "text/tab-separated-values;charset=utf-8");
    showToast("Cleaned dataset exported as TSV.");
  }

  function exportTreatmentCsv() {
    if (!state.results || !state.results.treatments || !state.results.treatments.length) {
      showToast("No treatment results to export.", "error");
      return;
    }

    const treatmentsStats = state.results.treatments;
    const header = [
      "Treatment",
      "IsControl",
      "Plots",
      "ReplicateCount",
      "AvgYield",
      "AvgDeltaYield",
      "AvgVariableCost",
      "AvgCapitalCost",
      "PVBenefits",
      "PVCosts",
      "NPV",
      "BCR",
      "ROI",
      "Rank"
    ];
    const lines = [header.join(",")];

    treatmentsStats.forEach((st) => {
      const row = [
        '"' + st.treatment.replace(/"/g, '""') + '"',
        st.isControl ? "1" : "0",
        st.plots,
        st.replicateCount,
        st.avgYield != null ? st.avgYield.toFixed(4) : "",
        st.avgDeltaYield != null ? st.avgDeltaYield.toFixed(4) : "",
        st.avgVarCost != null ? st.avgVarCost.toFixed(2) : "",
        st.avgCapCost != null ? st.avgCapCost.toFixed(2) : "",
        st.pvBenefits != null ? st.pvBenefits.toFixed(2) : "",
        st.pvCosts != null ? st.pvCosts.toFixed(2) : "",
        st.npv != null ? st.npv.toFixed(2) : "",
        st.bcr != null ? st.bcr.toFixed(4) : "",
        st.roi != null ? st.roi.toFixed(4) : "",
        st.rank
      ];
      lines.push(row.join(","));
    });

    const content = lines.join("\n");
    downloadFile("farming_cba_treatment_results.csv", content, "text/csv;charset=utf-8");
    showToast("Treatment results exported as CSV.");
  }

  function exportSensitivityCsv() {
    if (!state.sensitivity || !state.sensitivity.records) {
      showToast("No sensitivity records to export.", "error");
      return;
    }

    const header = [
      "DiscountRate",
      "PriceMultiplier",
      "BestTreatment",
      "BestNPV",
      "ControlNPV",
      "DifferenceBestMinusControl"
    ];
    const lines = [header.join(",")];

    state.sensitivity.records.forEach((rec) => {
      const delta = rec.bestNpv - rec.controlNpv;
      const row = [
        rec.discount,
        rec.priceMultiplier,
        '"' + rec.bestTreatment.replace(/"/g, '""') + '"',
        rec.bestNpv.toFixed(2),
        rec.controlNpv.toFixed(2),
        delta.toFixed(2)
      ];
      lines.push(row.join(","));
    });

    const content = lines.join("\n");
    downloadFile("farming_cba_sensitivity.csv", content, "text/csv;charset=utf-8");
    showToast("Sensitivity grid exported as CSV.");
  }

  function exportExcelWorkbook() {
    if (typeof XLSX === "undefined") {
      showToast("Excel library is not available in this browser.", "error");
      return;
    }

    const wb = XLSX.utils.book_new();

    // Raw data sheet
    if (state.header.length && state.rows.length) {
      const dataArray = [state.header];
      state.rows.forEach((row) => {
        dataArray.push(state.header.map((h) => row[h]));
      });
      const wsData = XLSX.utils.aoa_to_sheet(dataArray);
      XLSX.utils.book_append_sheet(wb, wsData, "RawData");
    }

    // Treatment results sheet
    if (state.results && state.results.treatments && state.results.treatments.length) {
      const t = state.results.treatments;
      const header = [
        "Treatment",
        "IsControl",
        "Plots",
        "ReplicateCount",
        "AvgYield",
        "AvgDeltaYield",
        "AvgVariableCost",
        "AvgCapitalCost",
        "PVBenefits",
        "PVCosts",
        "NPV",
        "BCR",
        "ROI",
        "Rank"
      ];
      const rows = [header];
      t.forEach((st) => {
        rows.push([
          st.treatment,
          st.isControl ? 1 : 0,
          st.plots,
          st.replicateCount,
          st.avgYield,
          st.avgDeltaYield,
          st.avgVarCost,
          st.avgCapCost,
          st.pvBenefits,
          st.pvCosts,
          st.npv,
          st.bcr,
          st.roi,
          st.rank
        ]);
      });
      const wsTreat = XLSX.utils.aoa_to_sheet(rows);
      XLSX.utils.book_append_sheet(wb, wsTreat, "Results");
    }

    // Sensitivity sheet
    if (state.sensitivity && state.sensitivity.records) {
      const header = [
        "DiscountRate",
        "PriceMultiplier",
        "BestTreatment",
        "BestNPV",
        "ControlNPV",
        "DifferenceBestMinusControl"
      ];
      const rows = [header];
      state.sensitivity.records.forEach((rec) => {
        const delta = rec.bestNpv - rec.controlNpv;
        rows.push([
          rec.discount,
          rec.priceMultiplier,
          rec.bestTreatment,
          rec.bestNpv,
          rec.controlNpv,
          delta
        ]);
      });
      const wsSens = XLSX.utils.aoa_to_sheet(rows);
      XLSX.utils.book_append_sheet(wb, wsSens, "Sensitivity");
    }

    XLSX.writeFile(wb, "farming_cba_workbook.xlsx");
    showToast("Excel workbook exported.");
  }

  // ======================
  // Scenarios
  // ======================

  function loadScenariosFromStorage() {
    try {
      const raw = localStorage.getItem(SCENARIO_KEY);
      if (!raw) {
        state.scenarios = [];
        renderScenarioSelect();
        return;
      }
      const parsed = JSON.parse(raw);
      state.scenarios = Array.isArray(parsed) ? parsed : [];
      renderScenarioSelect();
    } catch (_) {
      state.scenarios = [];
      renderScenarioSelect();
    }
  }

  function saveScenariosToStorage() {
    try {
      localStorage.setItem(SCENARIO_KEY, JSON.stringify(state.scenarios));
    } catch (_) {
      // ignore
    }
  }

  function renderScenarioSelect() {
    const select = document.getElementById("scenarioSelect");
    if (!select) return;
    select.innerHTML = "";

    if (!state.scenarios.length) {
      const opt = document.createElement("option");
      opt.value = "";
      opt.textContent = "No scenarios saved yet";
      select.appendChild(opt);
      return;
    }

    state.scenarios.forEach((sc, idx) => {
      const opt = document.createElement("option");
      opt.value = String(idx);
      opt.textContent = sc.name;
      select.appendChild(opt);
    });
  }

  function handleSaveScenario() {
    const nameInput = document.getElementById("scenarioName");
    if (!nameInput) return;
    const name = nameInput.value.trim();
    if (!name) {
      showToast("Please enter a scenario name.", "error");
      return;
    }

    const cfg = { ...state.config };
    const rawText = state.rawText;
    const scenario = { name, cfg, rawText };

    // Replace if same name
    const existingIndex = state.scenarios.findIndex((s) => s.name === name);
    if (existingIndex >= 0) {
      state.scenarios[existingIndex] = scenario;
    } else {
      state.scenarios.push(scenario);
    }

    saveScenariosToStorage();
    renderScenarioSelect();
    showToast("Scenario saved.");
  }

  function handleLoadScenario() {
    const select = document.getElementById("scenarioSelect");
    if (!select) return;
    const idx = parseInt(select.value, 10);
    if (isNaN(idx) || idx < 0 || idx >= state.scenarios.length) {
      showToast("Please select a saved scenario.", "error");
      return;
    }

    const scenario = state.scenarios[idx];
    state.config = { ...state.config, ...scenario.cfg };
    commitDataset(scenario.rawText, scenario.name);

    // Update config inputs
    syncConfigInputsFromState();

    showToast("Scenario loaded: " + scenario.name);
  }

  function handleDeleteScenario() {
    const select = document.getElementById("scenarioSelect");
    if (!select) return;
    const idx = parseInt(select.value, 10);
    if (isNaN(idx) || idx < 0 || idx >= state.scenarios.length) {
      showToast("Please select a saved scenario to delete.", "error");
      return;
    }

    const name = state.scenarios[idx].name;
    state.scenarios.splice(idx, 1);
    saveScenariosToStorage();
    renderScenarioSelect();
    showToast("Scenario deleted: " + name);
  }

  // ======================
  // Config
  // ======================

  function syncConfigInputsFromState() {
    const cfg = state.config;
    const map = {
      pricePerUnit: "pricePerUnit",
      horizonYears: "horizonYears",
      discountRate: "discountRate",
      discountLow: "discountLow",
      discountBase: "discountBase",
      discountHigh: "discountHigh",
      priceLow: "priceLow",
      priceBase: "priceBase",
      priceHigh: "priceHigh"
    };

    Object.entries(map).forEach(([key, id]) => {
      const el = document.getElementById(id);
      if (!el) return;
      if (key === "discountRate" || key.startsWith("discount")) {
        el.value = (cfg[key] * 100).toFixed(1).replace(/\.0$/, "");
      } else {
        el.value = String(cfg[key]);
      }
    });
  }

  function applyConfigFromInputs() {
    const cfg = { ...state.config };

    const priceInput = document.getElementById("pricePerUnit");
    const horizonInput = document.getElementById("horizonYears");
    const drInput = document.getElementById("discountRate");
    const dLow = document.getElementById("discountLow");
    const dBase = document.getElementById("discountBase");
    const dHigh = document.getElementById("discountHigh");
    const pLow = document.getElementById("priceLow");
    const pBase = document.getElementById("priceBase");
    const pHigh = document.getElementById("priceHigh");

    if (priceInput) {
      const v = parseFloat(priceInput.value);
      if (!isNaN(v) && v >= 0) cfg.pricePerUnit = v;
    }

    if (horizonInput) {
      const v = parseInt(horizonInput.value, 10);
      if (!isNaN(v) && v > 0) cfg.horizonYears = v;
    }

    if (drInput) {
      const v = parseFloat(drInput.value);
      if (!isNaN(v) && v >= 0) cfg.discountRate = v / 100;
    }

    if (dLow) {
      const v = parseFloat(dLow.value);
      if (!isNaN(v) && v >= 0) cfg.discountLow = v / 100;
    }
    if (dBase) {
      const v = parseFloat(dBase.value);
      if (!isNaN(v) && v >= 0) cfg.discountBase = v / 100;
    }
    if (dHigh) {
      const v = parseFloat(dHigh.value);
      if (!isNaN(v) && v >= 0) cfg.discountHigh = v / 100;
    }

    if (pLow) {
      const v = parseFloat(pLow.value);
      if (!isNaN(v) && v >= 0) cfg.priceLow = v;
    }
    if (pBase) {
      const v = parseFloat(pBase.value);
      if (!isNaN(v) && v >= 0) cfg.priceBase = v;
    }
    if (pHigh) {
      const v = parseFloat(pHigh.value);
      if (!isNaN(v) && v >= 0) cfg.priceHigh = v;
    }

    state.config = cfg;
    runAllCalculations();
    showToast("Economic settings applied.");
  }

  function resetConfigToDefault() {
    state.config = {
      pricePerUnit: 600,
      horizonYears: 10,
      discountRate: 0.06,
      discountLow: 0.04,
      discountBase: 0.06,
      discountHigh: 0.08,
      priceLow: 0.8,
      priceBase: 1.0,
      priceHigh: 1.2
    };
    syncConfigInputsFromState();
    runAllCalculations();
    showToast("Economic settings reset to default.");
  }

  // ======================
  // Appendix window
  // ======================

  function openAppendixWindow() {
    const win = window.open("", "_blank", "noopener,noreferrer");
    if (!win) return;

    const appendix = document.getElementById("tab-appendix");
    const content = appendix ? appendix.innerHTML : "<p>Appendix not available.</p>";

    win.document.write(
      "<!doctype html><html><head><meta charset='utf-8'><title>Technical appendix</title>" +
        "<style>body{font-family:system-ui,-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,'Helvetica Neue',Arial,sans-serif;padding:20px;max-width:800px;margin:0 auto;background:#f7f9fc;color:#111827;}h1,h2,h3{color:#004488;}p{line-height:1.6;}</style>" +
        "</head><body><h1>Technical appendix</h1>" +
        content +
        "</body></html>"
    );
    win.document.close();
  }

  // ======================
  // Copy helpers
  // ======================

  function copyFromElement(id) {
    const el = document.getElementById(id);
    if (!el) return;
    el.select();
    el.setSelectionRange(0, 999999);
    const ok = document.execCommand("copy");
    if (ok) {
      showToast("Copied to clipboard.");
    } else {
      showToast("Could not copy to clipboard.", "error");
    }
  }

  // ======================
  // Default dataset
  // ======================

  function loadDefaultDataset() {
    const el = document.getElementById("defaultDataset");
    if (!el) return;
    const text = el.value || el.textContent || "";
    if (!text.trim()) return;
    commitDataset(text, "Default trial");
  }

  // ======================
  // Event wiring
  // ======================

  function bindEvents() {
    // Data tab
    const fileInput = document.getElementById("fileInput");
    if (fileInput) {
      fileInput.addEventListener("change", (e) => {
        const file = e.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (evt) => {
          const text = evt.target.result;
          commitDataset(text, file.name);
        };
        reader.readAsText(file);
      });
    }

    const btnParsePaste = document.getElementById("btnParsePaste");
    if (btnParsePaste) {
      btnParsePaste.addEventListener("click", () => {
        const ta = document.getElementById("pasteInput");
        if (!ta) return;
        const text = ta.value.trim();
        if (!text) {
          showToast("Paste area is empty.", "error");
          return;
        }
        commitDataset(text, "Pasted data");
      });
    }

    const btnUseDefault = document.getElementById("btnUseDefault");
    if (btnUseDefault) {
      btnUseDefault.addEventListener("click", () => {
        loadDefaultDataset();
      });
    }

    // Config
    const btnApplyConfig = document.getElementById("btnApplyConfig");
    if (btnApplyConfig) {
      btnApplyConfig.addEventListener("click", applyConfigFromInputs);
    }

    const btnResetConfig = document.getElementById("btnResetConfig");
    if (btnResetConfig) {
      btnResetConfig.addEventListener("click", resetConfigToDefault);
    }

    // Scenarios
    const btnSaveScenario = document.getElementById("btnSaveScenario");
    if (btnSaveScenario) {
      btnSaveScenario.addEventListener("click", handleSaveScenario);
    }

    const btnLoadScenario = document.getElementById("btnLoadScenario");
    if (btnLoadScenario) {
      btnLoadScenario.addEventListener("click", handleLoadScenario);
    }

    const btnDeleteScenario = document.getElementById("btnDeleteScenario");
    if (btnDeleteScenario) {
      btnDeleteScenario.addEventListener("click", handleDeleteScenario);
    }

    // Filters
    const btnAll = document.getElementById("filterShowAll");
    if (btnAll) {
      btnAll.addEventListener("click", () => {
        if (state.results && state.results.treatments) {
          renderResultsLeaderboard(state.results.treatments, "all");
          showToast("Showing all treatments.");
        }
      });
    }

    const btnTopNpv = document.getElementById("filterTopNPV");
    if (btnTopNpv) {
      btnTopNpv.addEventListener("click", () => {
        if (state.results && state.results.treatments) {
          renderResultsLeaderboard(state.results.treatments, "topNPV");
          showToast("Showing top treatments by net present value.");
        }
      });
    }

    const btnTopBcr = document.getElementById("filterTopBCR");
    if (btnTopBcr) {
      btnTopBcr.addEventListener("click", () => {
        if (state.results && state.results.treatments) {
          renderResultsLeaderboard(state.results.treatments, "topBCR");
          showToast("Showing top treatments by benefit to cost ratio.");
        }
      });
    }

    const btnOnlyBetter = document.getElementById("filterOnlyBetter");
    if (btnOnlyBetter) {
      btnOnlyBetter.addEventListener("click", () => {
        if (state.results && state.results.treatments) {
          renderResultsLeaderboard(state.results.treatments, "onlyBetter");
          showToast("Showing treatments that beat the control.");
        }
      });
    }

    // Exports
    const btnExportClean = document.getElementById("btnExportClean");
    if (btnExportClean) {
      btnExportClean.addEventListener("click", exportCleanTsv);
    }

    const btnExportTreat = document.getElementById("btnExportTreatments");
    if (btnExportTreat) {
      btnExportTreat.addEventListener("click", exportTreatmentCsv);
    }

    const btnExportSens = document.getElementById("btnExportSensitivity");
    if (btnExportSens) {
      btnExportSens.addEventListener("click", exportSensitivityCsv);
    }

    const btnExportExcel = document.getElementById("btnExportExcel");
    if (btnExportExcel) {
      btnExportExcel.addEventListener("click", exportExcelWorkbook);
    }

    // AI copy buttons
    const btnCopyAi = document.getElementById("btnCopyAiBriefing");
    if (btnCopyAi) {
      btnCopyAi.addEventListener("click", () => copyFromElement("aiBriefingText"));
    }

    const btnCopyJson = document.getElementById("btnCopyResultsJson");
    if (btnCopyJson) {
      btnCopyJson.addEventListener("click", () => copyFromElement("resultsJson"));
    }

    // Appendix window
    const btnAppendixWindow = document.getElementById("btnOpenAppendixWindow");
    if (btnAppendixWindow) {
      btnAppendixWindow.addEventListener("click", openAppendixWindow);
    }
  }

  // ======================
  // Init
  // ======================

  document.addEventListener("DOMContentLoaded", () => {
    initTabs();
    loadScenariosFromStorage();
    syncConfigInputsFromState();
    loadDefaultDataset();
    bindEvents();
  });
})();
