const state = {
  taskId: "",
  taskName: "",
  chart1File: null,
  chart2File: null,
  ocrTestFile: null,
  ocrTesting: false,
  adminUnlocked: false,
  adminPassword: "",
  imageChecks: { chart1: null, chart2: null },
  started: false,
  loading: false,
  uploadMode: "ocr",
  masterRows: [],
  reviewRows: [],
  candidateRows: [],
  reviewDecisions: {},
  candidateDecisions: {},
  selectedReviewIndex: 0,
  selectedCandidateIndex: 0,
  chartView: "graph",
  chartIntent: "presentation",
  chartDirection: "down",
  chartStyle: "mono",
  chartMode: "a4",
  chartDepth: "all",
  showGroupRoot: false,
  hybridMode: "auto",
  hybridThreshold: 8,
  chartFontScale: 100,
  chartDensity: "full",
  selectedBranchId: "__all__",
  chartScale: 1,
  chartPanX: 0,
  chartPanY: 0,
  externalLayoutNeedsFit: false,
  printOrientation: "landscape",
  printTitle: "",
  printFitToPage: true,
  printScale: 100,
  printMargin: "normal",
  printFontSize: "medium",
  printSpacing: "normal",
  printForceOnePage: false,
  activityEvents: [],
  activityKeys: new Set(),
  selectedResultNodeIds: new Set(),
  chartShareholders: [],
  chartExternalEntities: [],
  undoStack: [],
  redoStack: [],
  hasUnsavedEdits: false,
  autoDraftTimer: null,
  toolbarCollapsed: false,
  currentView: "upload",
};

const elements = {
  pageTitle: document.getElementById("pageTitle"),
  navButtons: [...document.querySelectorAll(".nav-btn")],
  views: [...document.querySelectorAll(".view")],
  main: document.querySelector(".main"),
  chart1Input: document.getElementById("chart1Input"),
  chart2Input: document.getElementById("chart2Input"),
  chart1Meta: document.getElementById("chart1Meta"),
  chart2Meta: document.getElementById("chart2Meta"),
  chart1Preview: document.getElementById("chart1Preview"),
  chart2Preview: document.getElementById("chart2Preview"),
  taskNameInput: document.getElementById("taskNameInput"),
  uploadModeButtons: [...document.querySelectorAll(".upload-mode-btn")],
  manualCreatePanel: document.getElementById("manualCreatePanel"),
  manualRootNameInput: document.getElementById("manualRootNameInput"),
  manualTemplateSelect: document.getElementById("manualTemplateSelect"),
  createManualTaskBtn: document.getElementById("createManualTaskBtn"),
  imagePrecheck: document.getElementById("imagePrecheck"),
  taskStatusLine: document.getElementById("taskStatusLine"),
  startAnalysisBtn: document.getElementById("startAnalysisBtn"),
  exportBtn: document.getElementById("exportBtn"),
  metricsGrid: document.getElementById("metricsGrid"),
  overviewWarnings: document.getElementById("overviewWarnings"),
  reviewListTitle: document.getElementById("reviewListTitle"),
  reviewConfirmAllBtn: document.getElementById("reviewConfirmAllBtn"),
  reviewList: document.getElementById("reviewList"),
  reviewDetail: document.getElementById("reviewDetail"),
  candidateListTitle: document.getElementById("candidateListTitle"),
  candidateList: document.getElementById("candidateList"),
  candidateDetail: document.getElementById("candidateDetail"),
  searchInput: document.getElementById("searchInput"),
  statusFilter: document.getElementById("statusFilter"),
  resultTableTitle: document.getElementById("resultTableTitle"),
  resultTableBody: document.getElementById("resultTableBody"),
  undoBtn: document.getElementById("undoBtn"),
  redoBtn: document.getElementById("redoBtn"),
  saveDraftBtn: document.getElementById("saveDraftBtn"),
  restoreDraftBtn: document.getElementById("restoreDraftBtn"),
  draftStatusText: document.getElementById("draftStatusText"),
  resultSelectAll: document.getElementById("resultSelectAll"),
  batchFieldSelect: document.getElementById("batchFieldSelect"),
  batchValueInput: document.getElementById("batchValueInput"),
  batchApplyBtn: document.getElementById("batchApplyBtn"),
  batchDeleteBtn: document.getElementById("batchDeleteBtn"),
  selectedRowsHint: document.getElementById("selectedRowsHint"),
  taskSearchInput: document.getElementById("taskSearchInput"),
  taskStatusSelect: document.getElementById("taskStatusSelect"),
  taskSearchBtn: document.getElementById("taskSearchBtn"),
  taskCenterBody: document.getElementById("taskCenterBody"),
  addCompanyBtn: document.getElementById("addCompanyBtn"),
  bulkAddBtn: document.getElementById("bulkAddBtn"),
  addCompanyPanel: document.getElementById("addCompanyPanel"),
  bulkAddPanel: document.getElementById("bulkAddPanel"),
  addCompanyName: document.getElementById("addCompanyName"),
  addCompanyParent: document.getElementById("addCompanyParent"),
  addCompanyShare: document.getElementById("addCompanyShare"),
  bulkAddInput: document.getElementById("bulkAddInput"),
  cancelAddCompanyBtn: document.getElementById("cancelAddCompanyBtn"),
  saveAddCompanyBtn: document.getElementById("saveAddCompanyBtn"),
  cancelBulkAddBtn: document.getElementById("cancelBulkAddBtn"),
  saveBulkAddBtn: document.getElementById("saveBulkAddBtn"),
  ocrProviderSelect: document.getElementById("ocrProviderSelect"),
  ocrTestInput: document.getElementById("ocrTestInput"),
  ocrTestMeta: document.getElementById("ocrTestMeta"),
  ocrTestPreview: document.getElementById("ocrTestPreview"),
  ocrTestBtn: document.getElementById("ocrTestBtn"),
  ocrTestNote: document.getElementById("ocrTestNote"),
  ocrPromptProfileSelect: document.getElementById("ocrPromptProfileSelect"),
  ocrTestResult: document.getElementById("ocrTestResult"),
  adminLoginPanel: document.getElementById("adminLoginPanel"),
  adminPasswordInput: document.getElementById("adminPasswordInput"),
  adminUnlockBtn: document.getElementById("adminUnlockBtn"),
  ocrRefreshHistoryBtn: document.getElementById("ocrRefreshHistoryBtn"),
  ocrTestHistory: document.getElementById("ocrTestHistory"),
  chartViewButtons: [...document.querySelectorAll(".chart-view-btn")],
  chartIntentButtons: [...document.querySelectorAll(".chart-intent-btn")],
  chartDirectionButtons: [...document.querySelectorAll(".chart-direction-btn")],
  chartStyleButtons: [...document.querySelectorAll(".chart-style-btn")],
  chartModeButtons: [...document.querySelectorAll(".chart-mode-btn")],
  chartDepthButtons: [...document.querySelectorAll(".chart-depth-btn")],
  chartBranchPicker: document.getElementById("chartBranchPicker"),
  chartBranchSelect: document.getElementById("chartBranchSelect"),
  hybridModeSelect: document.getElementById("hybridModeSelect"),
  hybridThresholdSelect: document.getElementById("hybridThresholdSelect"),
  chartFontScaleSelect: document.getElementById("chartFontScaleSelect"),
  chartDensitySelect: document.getElementById("chartDensitySelect"),
  chartZoomButtons: [...document.querySelectorAll(".chart-zoom-btn")],
  chartZoomLabel: document.getElementById("chartZoomLabel"),
  chartShowRootToggle: document.getElementById("chartShowRootToggle"),
  printChartTitle: document.getElementById("printChartTitle"),
  chartContainer: document.getElementById("chartContainer"),
  chartLayoutBadge: document.getElementById("chartLayoutBadge"),
  chartAdvancedPanel: document.getElementById("chartAdvancedPanel"),
  chartLegend: document.getElementById("chartLegend"),
  exportPngBtn: document.getElementById("exportPngBtn"),
  exportHtmlBtn: document.getElementById("exportHtmlBtn"),
  printChartBtn: document.getElementById("printChartBtn"),
  toggleToolbarBtn: document.getElementById("toggleToolbarBtn"),
  openShareholderModalBtn: document.getElementById("openShareholderModalBtn"),
  openExternalEntityModalBtn: document.getElementById("openExternalEntityModalBtn"),
  shareholderList: document.getElementById("shareholderList"),
  externalEntityList: document.getElementById("externalEntityList"),
};

const pageTitles = {
  upload: "上傳任務",
  taskCenter: "任務中心",
  overview: "總覽",
  review: "待確認",
  candidates: "圖二新增候選",
  results: "結果主表",
  ocrTest: "管理員登入",
  chart: "股權架構圖",
};

const API_BASE = (window.API_BASE || "").replace(/\/$/, "");
const CHART_ZOOM_MIN = 0.35;
const CHART_ZOOM_MAX = 2.5;
const TASK_SNAPSHOT_KEY = "equity-review-last-task";
const CHART_VIEW_PREFS_KEY = "equity-chart-view-prefs";
const MAX_CHART1_MB = 3;
const MAX_CHART1_LONG_EDGE = 9000;
const MAX_CHART2_CHUNKS = 9;
const HYBRID_COLUMN_THRESHOLD = 8;

function clampNumber(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

async function apiGet(url) {
  const response = await fetch(API_BASE + url);
  if (!response.ok) throw new Error(`GET ${url} failed`);
  return response.json();
}

async function apiPost(url, body, isForm = false) {
  const response = await fetch(API_BASE + url, {
    method: "POST",
    headers: isForm ? undefined : { "Content-Type": "application/json" },
    body: isForm ? body : JSON.stringify(body),
  });
  if (!response.ok) {
    const payload = await response.json().catch(() => ({}));
    throw new Error(payload.message || `POST ${url} 失敗（${response.status}）`);
  }
  return response.json();
}

function saveTaskSnapshot(task) {
  if (!task?.id) return;
  try {
    localStorage.setItem(TASK_SNAPSHOT_KEY, JSON.stringify(task));
  } catch (_) {
    // Browser storage is a best-effort recovery path.
  }
}

function loadTaskSnapshot(taskId) {
  try {
    const task = JSON.parse(localStorage.getItem(TASK_SNAPSHOT_KEY) || "null");
    return task?.id === taskId ? task : null;
  } catch (_) {
    return null;
  }
}

async function apiGetAdmin(url) {
  const response = await fetch(API_BASE + url, {
    headers: state.adminPassword ? { "X-Admin-Test-Password": state.adminPassword } : {},
  });
  if (!response.ok) {
    const payload = await response.json().catch(() => ({}));
    throw new Error(payload.message || `GET ${url} 失敗（${response.status}）`);
  }
  return response.json();
}

function statusText(row) {
  if (row.node_status === "enriched") return "已自動補完";
  if (row.node_status === "review_match") return "待人工確認";
  return "圖一獨有";
}

function issueText(issueType) {
  return (
    {
      review_match: "名稱或層級需確認",
      chart1_only: "圖一有、圖二未補到",
      chart2_only: "圖二有、圖一沒有",
    }[issueType] || issueType
  );
}

function getRootRow() {
  return state.masterRows.find((row) => !row.chart1_parent || Number(row.chart1_level) === 0) || null;
}

function getGroupName() {
  const root = getRootRow();
  return (state.taskName || root?.canonical_name || root?.chart1_name || "未命名集團").trim();
}

function getChartRows() {
  const maxLevel = state.chartDepth === "all" ? Infinity : Number(state.chartDepth);
  const levelRows = state.masterRows.filter((row) => (Number(row.chart1_level) || 0) <= maxLevel);
  if (state.showGroupRoot) return levelRows;
  const root = getRootRow();
  if (!root?.node_id) return levelRows;
  return levelRows.filter((row) => row.node_id !== root.node_id);
}

function getTopLevelCompanyRows() {
  return state.masterRows.filter((row) => (Number(row.chart1_level) || 0) === 1);
}

function makeShareholderId() {
  if (window.crypto?.randomUUID) return window.crypto.randomUUID();
  return `sh_${Date.now()}_${Math.random().toString(16).slice(2)}`;
}

function shareholderTypeText(type) {
  return type === "person" ? "個人" : "公司";
}

function chartRowsWithShareholders(rows) {
  const byId = {};
  rows.forEach((row) => { byId[row.node_id] = row; });
  const extraRows = [];
  state.chartShareholders.forEach((holder) => {
    const target = byId[holder.target_node_id];
    if (!target) return;
    const targetLevel = Number(target.chart1_level) || 1;
    extraRows.push({
      node_id: `SH_${holder.id}`,
      canonical_name: holder.name,
      chart1_name: holder.name,
      chart1_level: Math.max(0, targetLevel - 1),
      chart1_parent: "",
      actual_controller_share: "",
      role_label: shareholderTypeText(holder.type),
      chart_note: holder.note || "",
      node_status: "enriched",
      is_chart_shareholder: true,
      shareholder_type: holder.type || "company",
      shareholder_target: holder.target_node_id,
      shareholder_share: holder.share || "",
    });
  });

  const groupMap = new Map();
  (state.chartExternalEntities || []).forEach((entity) => {
    const name = String(entity.name || "").trim() || "集團外主體";
    const groupName = String(entity.group || "集團外架構").trim() || "集團外架構";
    if (!groupMap.has(groupName)) {
      const groupId = `EXG_${groupName.replace(/\s+/g, "_")}_${groupMap.size + 1}`;
      groupMap.set(groupName, groupId);
      const placementMode = entity.placement_mode === "fixed" ? "fixed" : "auto";
      const manualX = Number.isFinite(Number(entity.manual_x)) ? Number(entity.manual_x) : null;
      const manualY = Number.isFinite(Number(entity.manual_y)) ? Number(entity.manual_y) : null;
      const externalScale = normalizeExternalScale(entity.external_scale);
      extraRows.push({
        node_id: groupId,
        canonical_name: groupName,
        chart1_name: groupName,
        chart1_level: 0,
        chart1_parent: "",
        actual_controller_share: "",
        role_label: "外部群組",
        chart_note: "集團外主體",
        node_status: "enriched",
        is_external_group: true,
        external_group_name: groupName,
        external_placement_mode: placementMode,
        external_manual_x: manualX,
        external_manual_y: manualY,
        external_scale: externalScale,
      });
    }
    const groupId = groupMap.get(groupName);
    const levels = Math.max(2, Math.min(4, Number(entity.levels) || 2));
    const layerNames = normalizeExternalLayerNames(entity);
    const layerShares = normalizeExternalLayerShares(entity);
    let parentId = groupId;
    for (let i = 0; i < levels; i += 1) {
      const isLast = i === levels - 1;
      const nodeId = `EX_${entity.id}_${i + 1}`;
      const nodeName = String(layerNames[i] || "").trim() || `${name} 第${i + 1}層`;
      extraRows.push({
        node_id: nodeId,
        canonical_name: nodeName,
        chart1_name: nodeName,
        chart1_level: i + 1,
        chart1_parent: parentId,
        actual_controller_share: i >= 1 ? String(layerShares[i - 1] || "").trim() : "",
        role_label: `集團外第${i + 1}層`,
        chart_note: isLast ? (entity.note || "") : "",
        node_status: "enriched",
        is_external_entity: true,
        is_external_mid: !isLast,
        external_group: groupName,
        external_link_target: isLast ? (entity.target_node_id || "") : "",
        external_link_share: isLast ? (entity.share || "") : "",
        external_scale: normalizeExternalScale(entity.external_scale),
      });
      parentId = nodeId;
    }
  });

  return [...extraRows, ...rows];
}

function getChartDepthLabel() {
  if (state.chartDepth === "1") return "顯示到一級";
  if (state.chartDepth === "2") return "顯示到二級";
  return "顯示全部層級";
}

function recommendationText(action) {
  return (
    {
      confirm_match_or_reject: "確認是否同一家公司",
      check_if_chart2_missing_or_inactive: "確認圖二是否未收錄或非存續",
      check_if_chart1_missing_node: "確認是否為圖一漏抽或更深層節點",
      manual_name_review: "人工核對名稱與候選公司",
    }[action] || action
  );
}

function updateTaskBadge() {
  return;
}

function setView(viewName) {
  state.currentView = viewName;
  elements.navButtons.forEach((button) => {
    button.classList.toggle("active", button.dataset.view === viewName);
  });
  elements.views.forEach((view) => {
    view.classList.toggle("active", view.id === viewName);
  });
  elements.main?.classList.toggle("chart-mode", viewName === "chart");
  elements.pageTitle.textContent = pageTitles[viewName];
  applyWorkspaceModeUI();
  syncActivityPanelVisibility(viewName);
  if (viewName === "taskCenter") loadTaskCenter();
  if (viewName === "chart") setTimeout(renderChart, 50); // 等 DOM 顯示後再渲染
}

function saveChartViewPrefs() {
  try {
    localStorage.setItem(CHART_VIEW_PREFS_KEY, JSON.stringify({
      toolbarCollapsed: state.toolbarCollapsed,
    }));
  } catch (_) {
    // best effort
  }
}

function applyWorkspaceModeUI() {
  const inChartView = state.currentView === "chart";
  document.body.classList.toggle("toolbar-collapsed", inChartView && state.toolbarCollapsed);
  if (elements.toggleToolbarBtn) {
    elements.toggleToolbarBtn.textContent = state.toolbarCollapsed ? "顯示工具列" : "隱藏工具列";
  }
}

function loadChartViewPrefs() {
  try {
    const prefs = JSON.parse(localStorage.getItem(CHART_VIEW_PREFS_KEY) || "null");
    if (!prefs || typeof prefs !== "object") return;
    state.toolbarCollapsed = Boolean(prefs.toolbarCollapsed);
  } catch (_) {
    // ignore invalid local data
  }
}

function enableStartIfReady() {
  if (state.uploadMode !== "ocr") {
    elements.startAnalysisBtn.disabled = true;
    return;
  }
  const blockers = getUploadBlockers();
  elements.startAnalysisBtn.disabled = !(state.chart1File && state.chart2File) || state.loading || blockers.length > 0;
}

function applyUploadModeUI() {
  const isOcr = state.uploadMode === "ocr";
  elements.uploadModeButtons.forEach((btn) => {
    const active = btn.dataset.uploadMode === state.uploadMode;
    btn.classList.toggle("active", active);
  });
  document.querySelectorAll(".drop-grid, .image-precheck, .upload-tips").forEach((el) => {
    if (!el) return;
    el.style.display = isOcr ? "" : "none";
  });
  if (elements.startAnalysisBtn) {
    elements.startAnalysisBtn.style.display = isOcr ? "" : "none";
  }
  if (elements.manualCreatePanel) {
    elements.manualCreatePanel.classList.toggle("hidden", isOcr);
  }
  enableStartIfReady();
}

function setUploadMode(mode) {
  state.uploadMode = mode === "manual" ? "manual" : "ocr";
  document.getElementById("uploadError")?.remove();
  applyUploadModeUI();
}

function enableOcrTestIfReady() {
  if (!elements.ocrTestBtn) return;
  elements.ocrTestBtn.disabled = !state.adminUnlocked || !state.ocrTestFile || state.ocrTesting;
}

function setPreview(file, metaEl, imgEl, dzEl) {
  if (!file) return;
  metaEl.textContent = `${file.name} · ${(file.size / 1024 / 1024).toFixed(2)} MB`;
  const objectUrl = URL.createObjectURL(file);
  imgEl.src = objectUrl;
  if (dzEl) dzEl.classList.add("has-file");
}

function estimateChart2Chunks(width, height) {
  const scaledHeight = width > 900 ? Math.round(height * (900 / width)) : height;
  if (scaledHeight <= 3600) return 1;
  return Math.ceil((scaledHeight - 1500) / (1500 - 180)) + 1;
}

function gradeLabel(score) {
  if (score >= 78) return { label: "高", tone: "high" };
  if (score >= 55) return { label: "中", tone: "mid" };
  return { label: "低", tone: "low" };
}

function inspectImageFile(file, kind) {
  return new Promise((resolve) => {
    if (!file) return resolve(null);
    const img = new Image();
    const url = URL.createObjectURL(file);
    img.onload = () => {
      const width = img.naturalWidth || img.width;
      const height = img.naturalHeight || img.height;
      const canvas = document.createElement("canvas");
      const maxSide = 360;
      const scale = Math.min(1, maxSide / Math.max(width, height));
      canvas.width = Math.max(1, Math.round(width * scale));
      canvas.height = Math.max(1, Math.round(height * scale));
      const ctx = canvas.getContext("2d", { willReadFrequently: true });
      ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
      const data = ctx.getImageData(0, 0, canvas.width, canvas.height).data;
      let edgeSum = 0;
      let samples = 0;
      for (let y = 1; y < canvas.height - 1; y += 3) {
        for (let x = 1; x < canvas.width - 1; x += 3) {
          const i = (y * canvas.width + x) * 4;
          const left = i - 4;
          const right = i + 4;
          const up = i - canvas.width * 4;
          const down = i + canvas.width * 4;
          const g = (data[i] + data[i + 1] + data[i + 2]) / 3;
          const lap = Math.abs(
            g * 4
            - (data[left] + data[left + 1] + data[left + 2]) / 3
            - (data[right] + data[right + 1] + data[right + 2]) / 3
            - (data[up] + data[up + 1] + data[up + 2]) / 3
            - (data[down] + data[down + 1] + data[down + 2]) / 3
          );
          edgeSum += lap;
          samples += 1;
        }
      }
      const sharpness = samples ? edgeSum / samples : 0;
      const minSide = Math.min(width, height);
      const longRatio = height / Math.max(width, 1);
      const aspectRatio = Math.max(width / Math.max(height, 1), height / Math.max(width, 1));
      const chunks = kind === "chart2" ? estimateChart2Chunks(width, height) : 1;
      const notes = [];
      let score = 100;

      if (minSide < 800) {
        score -= 25;
        notes.push("解析度偏低");
      } else if (minSide < 1100) {
        score -= 12;
        notes.push("解析度普通");
      }

      if (sharpness < 9) {
        score -= 24;
        notes.push("文字邊緣可能偏模糊");
      } else if (sharpness < 14) {
        score -= 10;
        notes.push("清晰度普通");
      }

      if (kind === "chart2") {
        if (chunks >= 6) {
          score -= 45;
          notes.push(`長截圖約需切成 ${chunks} 段，漏讀機率較高`);
        } else if (chunks >= 4) {
          score -= 25;
          notes.push(`長截圖約需切成 ${chunks} 段，建議確認結果`);
        }
      } else if (aspectRatio >= 3.2) {
        score -= 35;
        notes.push("超寬圖，文字容易過小或漏讀");
      } else if (aspectRatio >= 2.4) {
        score -= 20;
        notes.push("版面比例較寬，建議確認圖一骨架");
      }

      URL.revokeObjectURL(url);
      resolve({
        kind,
        width,
        height,
        sizeMb: file.size / 1024 / 1024,
        sharpness,
        chunks,
        score: clampNumber(Math.round(score), 0, 100),
        notes,
        ...gradeLabel(score),
      });
    };
    img.onerror = () => {
      URL.revokeObjectURL(url);
      resolve({
        kind,
        width: 0,
        height: 0,
        sizeMb: file.size / 1024 / 1024,
        sharpness: 0,
        chunks: 0,
        score: 0,
        notes: ["圖片無法預覽"],
        label: "低",
        tone: "low",
      });
    };
    img.src = url;
  });
}

function renderImagePrecheck() {
  if (!elements.imagePrecheck) return;
  getUploadBlockers();
  const checks = [state.imageChecks.chart1, state.imageChecks.chart2].filter(Boolean);
  if (!checks.length) {
    elements.imagePrecheck.innerHTML = "";
    elements.imagePrecheck.classList.remove("active");
    return;
  }
  const cards = [
    ["chart1", "圖一圖片條件"],
    ["chart2", "圖二圖片條件"],
  ].map(([key, title]) => {
    const item = state.imageChecks[key];
    if (!item) {
      return `<article class="precheck-card waiting"><h4>${title}</h4><p>尚未選擇圖片</p></article>`;
    }
    const extra = key === "chart2" && item.chunks > 1 ? ` · 約 ${item.chunks} 段辨識` : "";
    const advice = item.notes.length ? item.notes.join("、") : "圖片條件適合辨識";
    const blockedTags = (item.block_reasons || []).map((reason) => `<small style="color:#b91c1c;display:block;">${reason}</small>`).join("");
    return `
      <article class="precheck-card ${item.tone}">
        <div class="precheck-head">
          <h4>${title}</h4>
          <span>條件 ${item.label}</span>
        </div>
        <p>${item.width} × ${item.height}px${extra}</p>
        <small>${advice}</small>
        ${blockedTags}
      </article>`;
  }).join("");
  elements.imagePrecheck.innerHTML = `
    <div class="precheck-title">圖片條件檢查</div>
    <div class="precheck-grid">${cards}</div>
    <p class="precheck-note">這裡只評估清晰度、尺寸與長圖程度；實際辨識成功率會在 AI 完成後再估算。</p>`;
  elements.imagePrecheck.classList.add("active");
}

async function updateImageCheck(kind, file) {
  state.imageChecks[kind] = null;
  renderImagePrecheck();
  const result = await inspectImageFile(file, kind);
  if (state[`${kind}File`] !== file) return;
  state.imageChecks[kind] = result;
  renderImagePrecheck();
  enableStartIfReady();
}

function getUploadBlockers() {
  const blockers = [];
  const c1 = state.imageChecks.chart1;
  const c2 = state.imageChecks.chart2;
  if (c1) {
    const c1LongEdge = Math.max(c1.width || 0, c1.height || 0);
    if ((c1.sizeMb || 0) > MAX_CHART1_MB) blockers.push(`圖一超過 ${MAX_CHART1_MB}MB`);
    if (c1LongEdge > MAX_CHART1_LONG_EDGE) blockers.push(`圖一最長邊超過 ${MAX_CHART1_LONG_EDGE}px`);
  }
  if (c2) {
    if ((c2.chunks || 0) > MAX_CHART2_CHUNKS) blockers.push(`圖二預估分塊 ${c2.chunks} 段，超過上限 ${MAX_CHART2_CHUNKS} 段`);
  }
  if (c1) c1.block_reasons = [];
  if (c2) c2.block_reasons = [];
  blockers.forEach((message) => {
    if (message.startsWith("圖一") && c1) c1.block_reasons.push(message);
    if (message.startsWith("圖二") && c2) c2.block_reasons.push(message);
  });
  return blockers;
}

function ocrDisplayText(item) {
  if (typeof item === "string") return item;
  return item?.text || item?.company || item?.c || "";
}

function renderOcrTestResult(payload, elapsedMs = 0) {
  if (!elements.ocrTestResult) return;
  const items = Array.isArray(payload.items) ? payload.items : [];
  const candidates = Array.isArray(payload.company_candidates) ? payload.company_candidates : [];
  const providerLabel = [payload.engine || payload.provider || "OCR", payload.model, payload.prompt_profile].filter(Boolean).join(" · ");
  const shownCandidates = candidates.map(ocrDisplayText).filter(Boolean).slice(0, 80);
  const shownItems = items.map(ocrDisplayText).filter(Boolean).slice(0, 80);
  const candidateHtml = shownCandidates.length
    ? `<div class="ocr-test-chip-list">${shownCandidates.map((text) => `<span>${svgEscape(text)}</span>`).join("")}</div>`
    : `<p class="ocr-test-muted">沒有抓到疑似公司名稱。</p>`;
  const itemHtml = shownItems.length
    ? `<ol class="ocr-test-list">${shownItems.map((text) => `<li>${svgEscape(text)}</li>`).join("")}</ol>`
    : `<p class="ocr-test-muted">沒有文字結果。</p>`;

  elements.ocrTestResult.className = "ocr-test-result";
  elements.ocrTestResult.innerHTML = `
    <div class="ocr-test-result-head">
      <div>
        <p class="eyebrow">OCR Result</p>
        <h3>${svgEscape(providerLabel)}</h3>
      </div>
      <span>${(elapsedMs / 1000).toFixed(1)} 秒</span>
    </div>
    <div class="ocr-test-metrics">
      <span>${Number(payload.text_count || items.length)} 筆文字</span>
      <span>${Number(payload.company_candidate_count || candidates.length)} 個公司名稱候選</span>
    </div>
    <h4>公司名稱候選</h4>
    ${candidateHtml}
    <h4>文字清單</h4>
    ${itemHtml}`;
}

function renderOcrTestError(message) {
  if (!elements.ocrTestResult) return;
  elements.ocrTestResult.className = "ocr-test-result error";
  elements.ocrTestResult.innerHTML = `
    <div class="ocr-test-result-head">
      <div>
        <p class="eyebrow">OCR Result</p>
        <h3>測試失敗</h3>
      </div>
    </div>
    <p class="ocr-test-muted">${svgEscape(message)}</p>`;
}

function setAdminUnlocked(unlocked) {
  state.adminUnlocked = unlocked;
  elements.adminLoginPanel?.classList.toggle("unlocked", unlocked);
  document.querySelector("#ocrTest .ocr-test-layout")?.classList.toggle("locked", !unlocked);
  document.querySelector("#ocrTest .ocr-test-history-card")?.classList.toggle("locked", !unlocked);
  enableOcrTestIfReady();
}

async function unlockAdminTest() {
  state.adminPassword = elements.adminPasswordInput?.value.trim() || "";
  try {
    await loadOcrTestHistory();
    setAdminUnlocked(true);
  } catch (error) {
    setAdminUnlocked(false);
    renderOcrTestError(error.message || "管理員登入失敗。");
  }
}

function formatElapsed(ms) {
  const value = Number(ms || 0);
  if (!value) return "—";
  return `${(value / 1000).toFixed(1)} 秒`;
}

function renderOcrTestHistory(records) {
  if (!elements.ocrTestHistory) return;
  if (!records.length) {
    elements.ocrTestHistory.className = "ocr-test-history-empty";
    elements.ocrTestHistory.textContent = "目前尚無測試紀錄。";
    return;
  }
  elements.ocrTestHistory.className = "ocr-test-history-list";
  elements.ocrTestHistory.innerHTML = records.map((record) => {
    const candidates = Array.isArray(record.company_candidates)
      ? record.company_candidates.map(ocrDisplayText).filter(Boolean).slice(0, 8)
      : [];
    return `
      <article class="ocr-history-item">
        <div class="ocr-history-main">
          <strong>${svgEscape(record.provider || "OCR")} ${record.model ? `· ${svgEscape(record.model)}` : ""} ${record.prompt_profile ? `· ${svgEscape(record.prompt_profile)}` : ""}</strong>
          <span>${svgEscape(record.filename || "未命名圖片")}</span>
          ${record.note ? `<small>${svgEscape(record.note)}</small>` : ""}
        </div>
        <div class="ocr-history-stats">
          <span>${formatElapsed(record.elapsed_ms)}</span>
          <span>${Number(record.text_count || 0)} 筆文字</span>
          <span>${Number(record.company_candidate_count || 0)} 公司候選</span>
        </div>
        ${candidates.length ? `<div class="ocr-history-candidates">${candidates.map((name) => `<span>${svgEscape(name)}</span>`).join("")}</div>` : ""}
      </article>`;
  }).join("");
}

async function loadOcrTestHistory() {
  const payload = await apiGetAdmin("/api/ocr/tests?limit=50");
  renderOcrTestHistory(payload.tests || []);
}

async function runOcrTest() {
  if (!state.ocrTestFile) {
    renderOcrTestError("請先選擇測試圖片。");
    return;
  }
  const provider = elements.ocrProviderSelect?.value || "zhipu_ocr";
  const promptProfile = elements.ocrPromptProfileSelect?.value || "default";
  const originalText = elements.ocrTestBtn?.textContent || "開始 OCR 測試";
  const startedAt = performance.now();
  try {
    state.ocrTesting = true;
    if (elements.ocrTestBtn) elements.ocrTestBtn.textContent = "測試中...";
    enableOcrTestIfReady();
    if (elements.ocrTestResult) {
      elements.ocrTestResult.className = "ocr-test-result loading";
      elements.ocrTestResult.textContent = "正在辨識圖片文字...";
    }
    const form = new FormData();
    form.append("image", state.ocrTestFile);
    form.append("provider", provider);
    form.append("prompt_profile", promptProfile);
    form.append("save", "1");
    form.append("admin_password", state.adminPassword);
    if (elements.ocrTestNote?.value.trim()) form.append("note", elements.ocrTestNote.value.trim());
    const payload = await apiPost(`/api/ocr/probe?provider=${encodeURIComponent(provider)}&prompt_profile=${encodeURIComponent(promptProfile)}`, form, true);
    renderOcrTestResult(payload, Math.round(performance.now() - startedAt));
    await loadOcrTestHistory().catch(() => {});
  } catch (error) {
    console.error(error);
    renderOcrTestError(error.message || "OCR 測試失敗。");
  } finally {
    state.ocrTesting = false;
    if (elements.ocrTestBtn) elements.ocrTestBtn.textContent = originalText;
    enableOcrTestIfReady();
  }
}

function normalizeCompanyName(name) {
  return String(name || "")
    .replace(/[（）()]/g, "")
    .replace(/有限责任公司|有限公司|股份有限公司|集团|集團|公司/g, "")
    .replace(/\s+/g, "")
    .trim();
}

const REGION_PREFIXES = [
  "北京市", "天津市", "上海市", "重庆市", "重慶市",
  "河北省", "山西省", "辽宁省", "遼寧省", "吉林省", "黑龙江省", "黑龍江省",
  "江苏省", "江蘇省", "浙江省", "安徽省", "福建省", "江西省", "山东省", "山東省",
  "河南省", "湖北省", "湖南省", "广东省", "廣東省", "海南省", "四川省", "贵州省",
  "貴州省", "云南省", "雲南省", "陕西省", "陝西省", "甘肃省", "甘肅省", "青海省",
  "台湾省", "台灣省", "内蒙古", "內蒙古", "广西", "廣西", "西藏", "宁夏", "寧夏",
  "新疆", "北京", "天津", "上海", "重庆", "重慶", "河北", "山西", "辽宁", "遼寧",
  "吉林", "黑龙江", "黑龍江", "江苏", "江蘇", "浙江", "安徽", "福建", "江西",
  "山东", "山東", "河南", "湖北", "湖南", "广东", "廣東", "海南", "四川", "贵州",
  "貴州", "云南", "雲南", "陕西", "陝西", "甘肃", "甘肅", "青海", "台湾", "台灣",
  "深圳市", "广州市", "廣州市", "南京市", "杭州市", "宁波市", "寧波市", "苏州市",
  "蘇州市", "无锡市", "無錫市", "常州市", "扬州市", "揚州市", "泰州市", "南通市",
  "盐城市", "鹽城市", "淮安市", "连云港市", "連雲港市", "徐州市", "福州市", "厦门市",
  "廈門市", "清远市", "清遠市", "永州市", "道县", "道縣",
  "深圳", "广州", "廣州", "南京", "杭州", "宁波", "寧波", "苏州", "蘇州", "无锡",
  "無錫", "常州", "扬州", "揚州", "泰州", "南通", "盐城", "鹽城", "淮安", "连云港",
  "連雲港", "徐州", "福州", "厦门", "廈門", "清远", "清遠", "永州",
].sort((a, b) => b.length - a.length);

function stripLeadingRegion(name) {
  let text = normalizeCompanyName(name);
  let changed = true;
  while (changed && text.length >= 4) {
    changed = false;
    for (const region of REGION_PREFIXES) {
      if (text.startsWith(region) && text.length - region.length >= 4) {
        text = text.slice(region.length);
        changed = true;
        break;
      }
    }
  }
  return text;
}

function companyCoreForPattern(name) {
  return stripLeadingRegion(name)
    .replace(/发展|發展|科技|实业|實業|贸易|貿易|产业|產業|投资|投資|管理|控股/g, "")
    .slice(0, 6);
}

function hasLooseNameOverlap(a, b) {
  const x = normalizeCompanyName(a);
  const y = normalizeCompanyName(b);
  if (!x || !y) return false;
  return x.includes(y) || y.includes(x) || (x.length >= 4 && y.length >= 4 && x.slice(0, 4) === y.slice(0, 4));
}

function evaluateRecognitionSuccess({ masterRows = state.masterRows, chart2Rows = [] } = {}) {
  const rows = masterRows || [];
  const companyCount = rows.length;
  const levels = rows.map((row) => Number(row.chart1_level) || 0);
  const maxLevel = levels.length ? Math.max(...levels) : 0;
  const ids = new Set(rows.map((row) => row.node_id).filter(Boolean));
  const rootCount = rows.filter((row) => !row.chart1_parent || Number(row.chart1_level) === 0).length;
  const orphanCount = rows.filter((row) => row.chart1_parent && !ids.has(row.chart1_parent)).length;
  const chart1Names = rows.map((row) => row.canonical_name || row.chart1_name).filter(Boolean);
  const chart2Names = (chart2Rows || []).map((row) => row.company || row.chart2_name || row.c || "").filter(Boolean);
  const overlapCount = chart2Names.length
    ? chart2Names.filter((name) => chart1Names.some((other) => hasLooseNameOverlap(name, other))).length
    : 0;
  const overlapRate = chart2Names.length ? overlapCount / chart2Names.length : null;
  const levelOneCount = rows.filter((row) => Number(row.chart1_level) === 1).length;
  const flatLevelOneRatio = companyCount ? levelOneCount / companyCount : 0;
  const prefixCounts = {};
  const coreCounts = {};
  chart1Names.forEach((name) => {
    const normalized = normalizeCompanyName(name);
    const prefix = normalized.slice(0, 4);
    if (prefix.length >= 4) prefixCounts[prefix] = (prefixCounts[prefix] || 0) + 1;
    const core = companyCoreForPattern(name);
    if (core.length >= 4) coreCounts[core] = (coreCounts[core] || 0) + 1;
  });
  const largestPrefixGroup = Math.max(0, ...Object.values(prefixCounts));
  const largestCoreGroup = Math.max(0, ...Object.values(coreCounts));

  const notes = [];
  let score = 100;
  if (companyCount > 70) {
    score -= 25;
    notes.push(`公司數 ${companyCount} 家，建議拆分範圍`);
  } else if (companyCount > 40) {
    score -= 12;
    notes.push(`公司數 ${companyCount} 家，建議優先檢查主表`);
  } else if (companyCount) {
    notes.push(`公司數 ${companyCount} 家`);
  }

  if (maxLevel >= 5) {
    score -= 22;
    notes.push(`最大層級 ${maxLevel} 層`);
  } else if (maxLevel >= 4) {
    score -= 10;
    notes.push(`最大層級 ${maxLevel} 層`);
  }

  if (rootCount !== 1 && companyCount) {
    score -= 18;
    notes.push(`頂層主體 ${rootCount} 個，建議確認圖一骨架`);
  }
  if (orphanCount) {
    score -= 16;
    notes.push(`${orphanCount} 筆父層關係需確認`);
  }
  if (overlapRate !== null) {
    if (overlapRate < 0.12) {
      score -= 30;
      notes.push("圖一與圖二公司名稱幾乎不相符");
    } else if (overlapRate < 0.3) {
      score -= 15;
      notes.push("圖一與圖二公司名稱重疊普通");
    } else {
      notes.push("圖一與圖二有合理重疊");
    }
  }
  if (largestPrefixGroup >= 18 && largestPrefixGroup / Math.max(companyCount, 1) > 0.7) {
    score -= 12;
    notes.push("公司名稱高度相似，建議確認是否有誤讀");
  }
  if (largestCoreGroup >= 15 && largestCoreGroup / Math.max(companyCount, 1) > 0.45) {
    const severePattern = overlapRate !== null && overlapRate < 0.12;
    score -= severePattern ? 30 : 14;
    notes.push("圖一出現大量同型公司名稱，可能是 AI 補全城市或地區名");
  }
  if (companyCount >= 30 && flatLevelOneRatio > 0.82 && overlapRate !== null && overlapRate < 0.12) {
    score -= 12;
    notes.push("圖一層級過於扁平，建議先確認骨架是否誤讀");
  }

  const finalScore = clampNumber(Math.round(score), 0, 100);
  return {
    score: finalScore,
    ...gradeLabel(finalScore),
    notes: notes.length ? notes : ["結果結構適合進入審核"],
  };
}

function recognitionSuccessHtml(result) {
  if (!result) return "";
  return `
    <div class="recognition-success ${result.tone}">
      <strong>辨識成功率預估：${result.label}</strong>
      <span>${result.notes.join("；")}</span>
    </div>`;
}

function chart2ProgressText(progress = {}) {
  const current = Number(progress.current_chunk || 0);
  const total = Number(progress.total_chunks || 0);
  const rows = Number(progress.deduped_count || progress.rows_so_far || 0);
  const failed = Array.isArray(progress.failed_chunks) ? progress.failed_chunks.length : 0;
  if (!total) return "圖二辨識準備中…";
  const base = `圖二辨識中：${current}/${total} 塊，已抓到 ${rows} 家`;
  return failed ? `${base}，${failed} 塊需人工確認` : base;
}

function nowTimeLabel() {
  return new Date().toLocaleTimeString("zh-TW", { hour: "2-digit", minute: "2-digit", second: "2-digit" });
}

function ensureActivityPanel() {
  let panel = document.getElementById("activityPanel");
  if (panel) return panel;
  panel = document.createElement("aside");
  panel.id = "activityPanel";
  panel.className = "activity-panel";
  panel.innerHTML = `
    <div class="activity-head">
      <div>
        <p class="activity-eyebrow">Live</p>
        <h3>專案進度</h3>
      </div>
      <span class="activity-dot"></span>
    </div>
    <div id="activityList" class="activity-list">
      <p class="activity-empty">等待任務開始。</p>
    </div>`;
  document.body.appendChild(panel);
  return panel;
}

function addActivityEvent(key, title, detail = "", tone = "info") {
  if (state.activityKeys.has(key)) {
    return;
  }
  state.activityKeys.add(key);
  state.activityEvents.unshift({
    key,
    title,
    detail,
    tone,
    time: nowTimeLabel(),
  });
  state.activityEvents = state.activityEvents.slice(0, 10);
  renderActivityPanel();
}

function renderActivityPanel() {
  ensureActivityPanel();
  const list = document.getElementById("activityList");
  if (!list) return;
  if (!state.activityEvents.length) {
    list.innerHTML = `<p class="activity-empty">等待任務開始。</p>`;
    return;
  }
  list.innerHTML = state.activityEvents.slice(0, 3).map((event) => `
    <div class="activity-item ${event.tone}">
      <span class="activity-time">${event.time}</span>
      <div>
        <strong>${event.title}</strong>
        ${event.detail ? `<p>${event.detail}</p>` : ""}
      </div>
    </div>`).join("");
}

function syncActivityPanelVisibility(viewName = null) {
  const panel = ensureActivityPanel();
  const activeView = viewName || document.querySelector(".view.active")?.id || "upload";
  panel.style.display = activeView === "upload" ? "grid" : "none";
}

function trackWorkspaceActivity(phase, opts = {}) {
  const task = opts.task || {};
  const progress = opts.progress || task.chart2_progress || {};
  if (phase === "uploading") {
    addActivityEvent("uploading", "圖片上傳中", "建立任務。", "active");
  } else if (phase === "processing") {
    addActivityEvent("chart1-processing", "圖一辨識中", opts.msg || "建立結構骨架。", "active");
  } else if (phase === "chart1_ready") {
    const count = opts.summary?.master_count ?? state.masterRows.length;
    addActivityEvent("chart1-ready", "圖一完成", `${count || "多"} 家骨架。`, "done");
  } else if (phase === "processing_chart2") {
    const current = Number(progress.current_chunk || 0);
    const total = Number(progress.total_chunks || 0);
    const rows = Number(progress.deduped_count || progress.rows_so_far || 0);
    if (total && current) {
      addActivityEvent(`chart2-${current}-${rows}`, `圖二 ${current}/${total}`, `${rows} 家`, "active");
    } else {
      addActivityEvent("chart2-start", "圖二開始", opts.msg || "拆分長截圖。", "active");
    }
  } else if (phase === "chart2_confirm") {
    const count = (task.chart2_raw || []).length;
    addActivityEvent("chart2-confirm", "圖二完成", `${count} 家，等待確認。`, "done");
  } else if (phase === "ready") {
    const summary = opts.summary || {};
    addActivityEvent("ready", "合併完成", `主表 ${summary.master_count ?? state.masterRows.length} 家。`, "done");
  } else if (phase === "chart2_error") {
    addActivityEvent("chart2-error", "圖二需重試", opts.error || "圖一已保存。", "warn");
  } else if (phase === "cancelled") {
    addActivityEvent("task-cancelled", "任務已取消", "已停止後續分塊辨識。", "warn");
  } else if (phase === "error") {
    addActivityEvent(`error-${state.activityEvents.length}`, "任務中斷", "請重新整理或重試。", "warn");
  }
}

function makeMetric(label, value, theme) {
  return `
    <article class="metric-card ${theme}">
      <p class="metric-label">${label}</p>
      <p class="metric-value">${value}</p>
    </article>
  `;
}

function showAnalysisBanner(task) {
  // 舊版使用大橫幅；現在改成上方精簡狀態列。
  document.getElementById("analysisBanner")?.remove();

  const warning = task.analysis_warning;
  const mode = task.analysis_mode || "unknown";

  let msg = "尚未建立任務", type = "";
  if (warning) {
    msg = `AI 辨識未成功，目前顯示示範資料｜${warning}`;
    type = "status-warn";
  } else if (task.status === "cancelled" || task.status === "cancel_requested") {
    msg = `任務已取消｜任務 ID：${task.id}`;
    type = "status-warn";
  } else if (task.status === "processing" || task.status === "processing_chart2" || task.status === "chart1_ready") {
    msg = `任務進行中｜任務 ID：${task.id}`;
  } else if (mode === "qwen_vl") {
    const count = (task.master_rows || []).length;
    msg = `AI 辨識完成（Qwen-VL）｜任務 ID：${task.id}${count ? `｜${count} 家公司` : ""}`;
    type = "status-ok";
  }

  if (!elements.taskStatusLine) return;
  elements.taskStatusLine.textContent = msg;
  elements.taskStatusLine.className = `topbar-status ${type}`;
}

function hydrateTask(task) {
  saveTaskSnapshot(task);
  state.taskId = task.id;
  state.taskName = task.name;
  state.masterRows = task.master_rows || [];
  state.reviewRows = (task.review_rows || []).filter((row) => row.issue_type !== "chart2_only");
  state.candidateRows = task.candidate_rows || [];
  state.chartShareholders = task.chart_shareholders || [];
  state.chartExternalEntities = normalizeExternalEntities(task.chart_external_entities || []);
  state.printTitle = task.chart_print_settings?.title || "";
  state.printFitToPage = task.chart_print_settings?.fit_to_page !== false;
  state.printOrientation = task.chart_print_settings?.orientation || state.printOrientation;
  state.printScale = Number(task.chart_print_settings?.scale || 100);
  state.printMargin = task.chart_print_settings?.margin || "normal";
  state.printFontSize = task.chart_print_settings?.font_size || "medium";
  state.printSpacing = task.chart_print_settings?.spacing || "normal";
  state.printForceOnePage = task.chart_print_settings?.force_one_page === true;
  state.reviewDecisions = task.review_decisions || {};
  state.candidateDecisions = task.candidate_decisions || {};
  state.selectedReviewIndex = 0;
  state.selectedCandidateIndex = 0;
  state.selectedResultNodeIds = new Set();
  state.undoStack = [];
  state.redoStack = [];
  state.hasUnsavedEdits = false;
  if (state.autoDraftTimer) clearTimeout(state.autoDraftTimer);
  const draftSavedAt = task?.draft?.saved_at;
  if (draftSavedAt) setDraftStatus(`草稿：${String(draftSavedAt).replace("T", " ").slice(0, 19)}`);
  else setDraftStatus("尚未暫存");
  state.started = true;
  updateTaskBadge();
  showAnalysisBanner(task);
  renderOverview(task.summary || {});
  renderReviewList();
  renderReviewDetail();
  renderCandidateList();
  renderCandidateDetail();
  renderResults();
  renderShareholderPanel();
  renderExternalEntityPanel();
  updateUndoRedoButtons();
}

function applyTaskRefresh(payload) {
  if (!payload) return;
  if (payload.master_rows) state.masterRows = payload.master_rows;
  if (payload.review_rows) state.reviewRows = payload.review_rows.filter((row) => row.issue_type !== "chart2_only");
  if (payload.candidate_rows) state.candidateRows = payload.candidate_rows;
  if (payload.chart_shareholders) state.chartShareholders = payload.chart_shareholders;
  if (payload.chart_external_entities) state.chartExternalEntities = normalizeExternalEntities(payload.chart_external_entities);
  if (payload.chart_print_settings) {
    state.printTitle = payload.chart_print_settings.title || "";
    state.printFitToPage = payload.chart_print_settings.fit_to_page !== false;
    state.printOrientation = payload.chart_print_settings.orientation || state.printOrientation;
    state.printScale = Number(payload.chart_print_settings.scale || 100);
    state.printMargin = payload.chart_print_settings.margin || "normal";
    state.printFontSize = payload.chart_print_settings.font_size || "medium";
    state.printSpacing = payload.chart_print_settings.spacing || "normal";
    state.printForceOnePage = payload.chart_print_settings.force_one_page === true;
  }
  state.selectedResultNodeIds = new Set(
    [...state.selectedResultNodeIds].filter((nodeId) => state.masterRows.some((row) => row.node_id === nodeId))
  );
  if (payload.review_decisions) state.reviewDecisions = payload.review_decisions;
  if (payload.candidate_decisions) state.candidateDecisions = payload.candidate_decisions;
  renderOverview(payload.summary || {});
  renderReviewList();
  renderReviewDetail();
  renderCandidateList();
  renderCandidateDetail();
  renderResults();
  renderShareholderPanel();
  renderExternalEntityPanel();
  updateUndoRedoButtons();
}

function normalizeExternalEntities(entities) {
  return (entities || []).map((entity) => ({
    ...entity,
    levels: Math.max(2, Math.min(4, Number(entity?.levels) || 2)),
    layer_names: normalizeExternalLayerNames(entity),
    layer_shares: normalizeExternalLayerShares(entity),
    external_scale: normalizeExternalScale(entity?.external_scale),
    placement_mode: entity?.placement_mode === "fixed" ? "fixed" : "auto",
    manual_x: Number.isFinite(Number(entity?.manual_x)) ? Number(entity.manual_x) : null,
    manual_y: Number.isFinite(Number(entity?.manual_y)) ? Number(entity.manual_y) : null,
  }));
}

function normalizeExternalScale(value) {
  const n = Number(value);
  if (!Number.isFinite(n)) return 100;
  return Math.max(70, Math.min(130, Math.round(n)));
}

function normalizeExternalLayerNames(entity = {}) {
  const levels = Math.max(2, Math.min(4, Number(entity?.levels) || 2));
  const seed = Array.isArray(entity.layer_names) ? entity.layer_names.map((v) => String(v || "").trim()) : [];
  while (seed.length < levels) seed.push("");
  seed.length = levels;
  if (!seed[0]) seed[0] = String(entity.name || "").trim() || "最上層";
  for (let i = 1; i < levels; i += 1) {
    if (!seed[i]) seed[i] = `${seed[i - 1]} 第${i + 1}層`;
  }
  return seed;
}

function normalizeExternalLayerShares(entity = {}) {
  const levels = Math.max(2, Math.min(4, Number(entity?.levels) || 2));
  const expected = Math.max(0, levels - 1);
  const seed = Array.isArray(entity.layer_shares) ? entity.layer_shares.map((v) => String(v || "").trim()) : [];
  while (seed.length < expected) seed.push("");
  seed.length = expected;
  return seed;
}

function snapshotTaskState() {
  return {
    master_rows: JSON.parse(JSON.stringify(state.masterRows || [])),
    review_rows: JSON.parse(JSON.stringify(state.reviewRows || [])),
    candidate_rows: JSON.parse(JSON.stringify(state.candidateRows || [])),
    review_decisions: JSON.parse(JSON.stringify(state.reviewDecisions || {})),
    candidate_decisions: JSON.parse(JSON.stringify(state.candidateDecisions || {})),
  };
}

function pushUndoSnapshot(snapshot) {
  state.undoStack.push(snapshot);
  if (state.undoStack.length > 20) state.undoStack.shift();
  state.redoStack = [];
  markTaskDirty();
  updateUndoRedoButtons();
}

function updateUndoRedoButtons() {
  if (elements.undoBtn) elements.undoBtn.disabled = !state.taskId || state.undoStack.length === 0;
  if (elements.redoBtn) elements.redoBtn.disabled = !state.taskId || state.redoStack.length === 0;
}

async function restoreTaskSnapshot(snapshot) {
  const payload = await apiPost(`/api/tasks/${state.taskId}/replace-state`, snapshot);
  applyTaskRefresh(payload);
}

async function undoTaskEdit() {
  if (!state.taskId || !state.undoStack.length) return;
  const current = snapshotTaskState();
  const target = state.undoStack.pop();
  state.redoStack.push(current);
  await restoreTaskSnapshot(target);
  updateUndoRedoButtons();
}

async function redoTaskEdit() {
  if (!state.taskId || !state.redoStack.length) return;
  const current = snapshotTaskState();
  const target = state.redoStack.pop();
  state.undoStack.push(current);
  await restoreTaskSnapshot(target);
  updateUndoRedoButtons();
}

function setDraftStatus(text, tone = "") {
  if (!elements.draftStatusText) return;
  elements.draftStatusText.textContent = text;
  elements.draftStatusText.className = `topbar-status ${tone}`.trim();
}

function markTaskDirty() {
  state.hasUnsavedEdits = true;
  setDraftStatus("尚有未暫存變更", "status-warn");
  if (state.autoDraftTimer) clearTimeout(state.autoDraftTimer);
  if (!state.taskId) return;
  state.autoDraftTimer = setTimeout(() => {
    saveDraftSnapshot(true).catch((error) => console.error("auto draft save failed", error));
  }, 1500);
}

async function saveDraftSnapshot(silent = false) {
  if (!state.taskId) return;
  const payload = { state: snapshotTaskState() };
  try {
    const result = await apiPost(`/api/tasks/${state.taskId}/save-draft`, payload);
    state.hasUnsavedEdits = false;
    setDraftStatus(`已暫存：${(result.saved_at || "").replace("T", " ").slice(0, 19)}`, "status-ok");
    if (!silent) alert("草稿已暫存。");
  } catch (error) {
    state.hasUnsavedEdits = true;
    setDraftStatus(`暫存失敗：${error.message}`, "status-warn");
    if (!silent) alert(`暫存失敗：${error.message}`);
    throw error;
  }
}

async function restoreDraftSnapshot() {
  if (!state.taskId) return;
  const payload = await apiPost(`/api/tasks/${state.taskId}/restore-draft`, {});
  applyTaskRefresh(payload);
  state.hasUnsavedEdits = false;
  setDraftStatus("已還原草稿", "status-ok");
}

async function loadTaskCenter() {
  if (!elements.taskCenterBody) return;
  const q = encodeURIComponent((elements.taskSearchInput?.value || "").trim());
  const status = encodeURIComponent((elements.taskStatusSelect?.value || "").trim());
  elements.taskCenterBody.innerHTML = `<tr><td colspan="7">讀取中…</td></tr>`;
  try {
    const payload = await apiGet(`/api/tasks?q=${q}&status=${status}&limit=200`);
    const rows = payload.tasks || [];
    if (!rows.length) {
      elements.taskCenterBody.innerHTML = `<tr><td colspan="7">找不到任務</td></tr>`;
      return;
    }
    elements.taskCenterBody.innerHTML = rows.map((task) => `
      <tr>
        <td><code>${task.id || ""}</code></td>
        <td>${task.name || "未命名"}</td>
        <td>${task.status || "-"}</td>
        <td>${Math.max((task.master_count || 0) - 1, 0)}</td>
        <td>${(task.created_at || "").replace("T", " ").slice(0, 19) || "-"}</td>
        <td>${(task.updated_at || "").replace("T", " ").slice(0, 19)}</td>
        <td>
          <button class="ghost-btn task-open-btn" data-task-id="${task.id}">開啟</button>
          <button class="ghost-btn task-clone-btn" data-task-id="${task.id}">複製</button>
          <button class="ghost-btn task-delete-btn" data-task-id="${task.id}">刪除</button>
        </td>
      </tr>
    `).join("");
    elements.taskCenterBody.querySelectorAll(".task-open-btn").forEach((btn) => {
      btn.addEventListener("click", () => openTaskFromCenter(btn.dataset.taskId));
    });
    elements.taskCenterBody.querySelectorAll(".task-clone-btn").forEach((btn) => {
      btn.addEventListener("click", () => cloneTaskFromCenter(btn.dataset.taskId));
    });
    elements.taskCenterBody.querySelectorAll(".task-delete-btn").forEach((btn) => {
      btn.addEventListener("click", () => deleteTaskFromCenter(btn.dataset.taskId));
    });
  } catch (err) {
    elements.taskCenterBody.innerHTML = `<tr><td colspan="7">讀取失敗：${err.message}</td></tr>`;
  }
}

async function deleteTaskFromCenter(taskId) {
  if (!taskId) return;
  if (!window.confirm(`確定刪除任務 ${taskId}？此操作無法復原。`)) return;
  const response = await fetch(`${API_BASE}/api/tasks/${taskId}`, { method: "DELETE" });
  if (!response.ok) {
    const payload = await response.json().catch(() => ({}));
    alert(payload.message || `刪除失敗（${response.status}）`);
    return;
  }
  await loadTaskCenter();
}

async function openTaskFromCenter(taskId) {
  if (!taskId) return;
  const task = await apiGet(`/api/tasks/${taskId}`);
  hydrateTask(task);
  if (task.status === "chart2_ocr_done") {
    renderWorkspace("chart2_confirm", { task });
    setView("upload");
    return;
  }
  if (task.status === "processing_chart2") {
    renderWorkspace("processing_chart2", { progress: task.chart2_progress, summary: task.summary });
  } else if (task.status === "chart2_error") {
    renderWorkspace("chart2_error", { taskId: task.id, error: task.error, summary: task.summary, diagnostics: task.analysis_diagnostics });
  } else if (task.status === "cancelled" || task.status === "cancel_requested") {
    renderWorkspace("cancelled", { summary: task.summary });
  } else {
    renderWorkspace("ready", { summary: task.summary || {} });
  }
  setView("overview");
}

async function cloneTaskFromCenter(taskId) {
  if (!taskId) return;
  await apiPost(`/api/tasks/${taskId}/clone`, {});
  await loadTaskCenter();
}

function updateResultSelectAllState(rows) {
  if (!elements.resultSelectAll) return;
  if (!rows.length) {
    elements.resultSelectAll.checked = false;
    elements.resultSelectAll.indeterminate = false;
    updateSelectedRowsHint(0);
    return;
  }
  const selected = rows.filter((row) => state.selectedResultNodeIds.has(row.node_id)).length;
  elements.resultSelectAll.checked = selected > 0 && selected === rows.length;
  elements.resultSelectAll.indeterminate = selected > 0 && selected < rows.length;
  updateSelectedRowsHint(selected);
}

function updateSelectedRowsHint(selected = state.selectedResultNodeIds.size) {
  if (!elements.selectedRowsHint) return;
  elements.selectedRowsHint.textContent = selected ? `已勾選 ${selected} 間，可批次編輯或拖曳` : "尚未勾選";
  if (elements.batchDeleteBtn) elements.batchDeleteBtn.disabled = !selected;
  if (elements.batchApplyBtn) elements.batchApplyBtn.disabled = !selected;
}

async function runBatchUpdate() {
  const nodeIds = [...state.selectedResultNodeIds];
  if (!state.taskId || !nodeIds.length) {
    alert("請先勾選至少一筆公司。");
    return;
  }
  const field = elements.batchFieldSelect?.value || "";
  const value = elements.batchValueInput?.value || "";
  const before = snapshotTaskState();
  const payload = await apiPost(`/api/tasks/${state.taskId}/batch-update`, { node_ids: nodeIds, field, value });
  applyTaskRefresh(payload);
  pushUndoSnapshot(before);
}

async function runBatchDelete() {
  const nodeIds = [...state.selectedResultNodeIds];
  if (!state.taskId || !nodeIds.length) {
    alert("請先勾選至少一筆公司。");
    return;
  }
  if (!confirm(`確定要刪除已勾選的 ${nodeIds.length} 間公司嗎？`)) return;
  const before = snapshotTaskState();
  const payload = await apiPost(`/api/tasks/${state.taskId}/batch-delete`, { node_ids: nodeIds });
  state.selectedResultNodeIds = new Set();
  applyTaskRefresh(payload);
  pushUndoSnapshot(before);
}

function renderOverview(summary) {
  const total = summary.master_count ?? state.masterRows.length;
  const enriched = summary.enriched_count ?? state.masterRows.filter((row) => row.node_status === "enriched").length;
  const pending = summary.review_count ?? state.reviewRows.length;
  const chart1Only = summary.chart1_only_count ?? state.masterRows.filter((row) => row.node_status === "chart1_only").length;
  const candidates = summary.candidate_count ?? state.candidateRows.length;

  elements.metricsGrid.innerHTML = [
    makeMetric("主表公司數", total, "blue"),
    makeMetric("已自動補完", enriched, "green"),
    makeMetric("待人工確認", pending, "gold"),
    makeMetric("圖一獨有", chart1Only, "slate"),
    makeMetric("圖二新增候選", candidates, "orange"),
  ].join("");

  const success = evaluateRecognitionSuccess();
  const warnings = [
    `辨識成功率預估：${success.label}。${success.notes.join("；")}。`,
    `${pending} 筆資料仍需要人工確認，請先處理這一區。`,
    `${candidates} 筆圖二新增候選尚未決定是否加入主表。`,
    "第二階段股權架構圖會依賴這裡的最終審核結果，所以名稱與上層公司要盡量確認乾淨。",
  ];
  elements.overviewWarnings.innerHTML = warnings.map((text) => `<li>${text}</li>`).join("");
}

function renderReviewList() {
  elements.reviewListTitle.textContent = `${state.reviewRows.length} 筆待確認`;
  if (elements.reviewConfirmAllBtn) {
    elements.reviewConfirmAllBtn.disabled = !state.taskId || state.reviewRows.length === 0;
    elements.reviewConfirmAllBtn.textContent = state.reviewRows.length ? `全部確認 ${state.reviewRows.length} 筆` : "全部確認";
  }
  elements.reviewList.innerHTML = state.reviewRows
    .map((row, index) => {
      const key = row.candidate_node_id || row.chart2_name;
      const decision = state.reviewDecisions[key];
      return `
        <article class="review-item ${index === state.selectedReviewIndex ? "active" : ""}" data-review-index="${index}">
          <h4>${row.chart1_name || row.chart2_name}</h4>
          <div class="pill-row">
            <span class="pill warning">${issueText(row.issue_type)}</span>
            <span class="pill info">分數 ${row.match_score || "—"}</span>
            ${decision?.decision ? `<span class="pill slate">已填：${decision.decision}</span>` : ""}
          </div>
        </article>
      `;
    })
    .join("");

  [...elements.reviewList.querySelectorAll("[data-review-index]")].forEach((item) => {
    item.addEventListener("click", () => {
      state.selectedReviewIndex = Number(item.dataset.reviewIndex);
      renderReviewList();
      renderReviewDetail();
    });
  });
}

async function confirmAllReviewRows() {
  if (!state.taskId || !state.reviewRows.length) return;
  if (!confirm(`確定要將 ${state.reviewRows.length} 筆待確認全部視為可用嗎？`)) return;
  const before = snapshotTaskState();
  const response = await apiPost(`/api/tasks/${state.taskId}/review-confirm-all`, {});
  applyTaskRefresh(response);
  pushUndoSnapshot(before);
  setView("results");
}

function renderReviewDetail() {
  const row = state.reviewRows[state.selectedReviewIndex];
  if (!row) {
    elements.reviewDetail.innerHTML = `
      <div class="detail-empty">
        <h3>目前沒有待確認項目</h3>
        <p>圖一辨識到的公司已可先進入結果主表檢查與調整。</p>
      </div>
    `;
    return;
  }
  const key = row.candidate_node_id || row.chart2_name;
  const saved = state.reviewDecisions[key] || {};

  elements.reviewDetail.innerHTML = `
    <div class="detail-section">
      <div class="section-head">
        <div>
          <p class="eyebrow">人工確認</p>
          <h3>${row.chart1_name || row.chart2_name}</h3>
        </div>
      </div>
      <div class="detail-grid">
        <div class="info-box"><span>圖一名稱</span>${row.chart1_name || "—"}</div>
        <div class="info-box"><span>圖二名稱</span>${row.chart2_name || "—"}</div>
        <div class="info-box"><span>問題類型</span>${issueText(row.issue_type)}</div>
        <div class="info-box"><span>建議處理方式</span>${recommendationText(row.recommended_action)}</div>
      </div>
      <div class="info-box">
        <span>系統說明</span>
        ${row.review_note || "—"}
      </div>
      <div class="detail-grid">
        <label class="field">
          <span>人工確認結果</span>
          <select id="reviewDecision">
            <option value="">請選擇</option>
            ${["確認一致", "不是同一家公司", "圖一漏節點", "暫不處理"]
              .map((option) => `<option value="${option}" ${saved.decision === option ? "selected" : ""}>${option}</option>`)
              .join("")}
          </select>
        </label>
        <label class="field">
          <span>人工修正公司名</span>
          <input id="reviewName" type="text" value="${saved.corrected_name || ""}" placeholder="如需修正公司名稱，填在這裡" />
        </label>
        <label class="field">
          <span>人工修正層級</span>
          <input id="reviewLevel" type="text" value="${saved.corrected_level || ""}" placeholder="例如：2 或 二级子公司" />
        </label>
        <label class="field">
          <span>人工修正上層公司</span>
          <input id="reviewParent" type="text" value="${saved.corrected_parent || ""}" placeholder="填上層公司名稱" />
        </label>
      </div>
      <label class="field">
        <span>人工備註</span>
        <textarea id="reviewNote" placeholder="只寫需要傳達給下一位使用者的關鍵說明">${saved.note || ""}</textarea>
      </label>
      <div class="detail-actions">
        <button id="reviewPrevBtn" class="ghost-btn">上一筆</button>
        <button id="reviewSaveBtn" class="primary-btn">儲存本筆</button>
        <button id="reviewNextBtn" class="ghost-btn">下一筆</button>
      </div>
    </div>
  `;

  document.getElementById("reviewSaveBtn").addEventListener("click", async () => {
    const payload = {
      task_id: state.taskId,
      key,
      decision: document.getElementById("reviewDecision").value,
      corrected_name: document.getElementById("reviewName").value,
      corrected_level: document.getElementById("reviewLevel").value,
      corrected_parent: document.getElementById("reviewParent").value,
      note: document.getElementById("reviewNote").value,
    };
    const response = await apiPost("/api/review-decision", payload);
    applyTaskRefresh(response);
  });
  document.getElementById("reviewPrevBtn").addEventListener("click", () => {
    state.selectedReviewIndex = Math.max(0, state.selectedReviewIndex - 1);
    renderReviewList();
    renderReviewDetail();
  });
  document.getElementById("reviewNextBtn").addEventListener("click", () => {
    state.selectedReviewIndex = Math.min(state.reviewRows.length - 1, state.selectedReviewIndex + 1);
    renderReviewList();
    renderReviewDetail();
  });
}

function renderCandidateList() {
  elements.candidateListTitle.textContent = `${state.candidateRows.length} 筆候選`;
  elements.candidateList.innerHTML = state.candidateRows
    .map((row, index) => {
      const decision = state.candidateDecisions[row.chart2_name];
      const isAdded = decision?.decision === "加入主表";
      return `
        <article class="review-item ${index === state.selectedCandidateIndex ? "active" : ""}" data-candidate-index="${index}">
          <h4 class="candidate-title">${row.chart2_name}</h4>
          <div class="pill-row">
            <span class="pill slate">${row.subsidiary_level_label || "未標級別"}</span>
            <span class="pill info">${row.company_status || "未標狀態"}</span>
            ${decision?.decision ? `<span class="pill ${isAdded ? "success" : "warning"}">${isAdded ? "已新增" : `已填：${decision.decision}`}</span>` : ""}
          </div>
        </article>
      `;
    })
    .join("");

  [...elements.candidateList.querySelectorAll("[data-candidate-index]")].forEach((item) => {
    item.addEventListener("click", () => {
      state.selectedCandidateIndex = Number(item.dataset.candidateIndex);
      renderCandidateList();
      renderCandidateDetail();
    });
  });
}

function renderCandidateDetail() {
  const row = state.candidateRows[state.selectedCandidateIndex];
  if (!row) return;
  const saved = state.candidateDecisions[row.chart2_name] || {};
  const parentOptions = state.masterRows
    .map((master) => `<option value="${master.chart1_name}" ${saved.parent === master.chart1_name ? "selected" : ""}>${master.chart1_name}</option>`)
    .join("");

  elements.candidateDetail.innerHTML = `
    <div class="detail-section">
      <div class="section-head">
        <div>
          <p class="eyebrow">新增候選</p>
          <h3>${row.chart2_name}</h3>
        </div>
      </div>
      <div class="detail-grid">
        <div class="info-box"><span>法人代表</span>${row.legal_representative || "—"}</div>
        <div class="info-box"><span>成立時間</span>${row.established_date || "—"}</div>
        <div class="info-box"><span>資本額</span>${row.registered_capital || "—"}</div>
        <div class="info-box"><span>實控人持股</span>${row.actual_controller_share || "—"}</div>
      </div>
      <div class="info-box">
        <span>未併入原因</span>
        ${row.reason_not_merged || "—"}
      </div>
      <div class="detail-grid">
        <label class="field">
          <span>是否加入主表</span>
          <select id="candidateDecision">
            <option value="">請選擇</option>
            ${["加入主表", "先不加入", "暫不處理"]
              .map((option) => `<option value="${option}" ${saved.decision === option ? "selected" : ""}>${option}</option>`)
              .join("")}
          </select>
        </label>
        <label class="field">
          <span>指定上層公司</span>
          <select id="candidateParent">
            <option value="">請選擇上層公司</option>
            ${parentOptions}
          </select>
        </label>
        <label class="field">
          <span>人工修正公司名</span>
          <input id="candidateName" type="text" value="${saved.corrected_name || ""}" placeholder="如需修正名稱，填在這裡" />
        </label>
      </div>
      <label class="field">
        <span>人工備註</span>
        <textarea id="candidateNote" placeholder="例如：確定是圖一未展開的四級子公司">${saved.note || ""}</textarea>
      </label>
      <div class="detail-actions">
        <button id="candidatePrevBtn" class="ghost-btn">上一筆</button>
        <button id="candidateSaveBtn" class="primary-btn">儲存本筆</button>
        <button id="candidateNextBtn" class="ghost-btn">下一筆</button>
      </div>
    </div>
  `;

  document.getElementById("candidateSaveBtn").addEventListener("click", async () => {
    const payload = {
      task_id: state.taskId,
      key: row.chart2_name,
      decision: document.getElementById("candidateDecision").value,
      parent: document.getElementById("candidateParent").value,
      corrected_name: document.getElementById("candidateName").value,
      note: document.getElementById("candidateNote").value,
    };
    const response = await apiPost("/api/candidate-decision", payload);
    if (state.selectedCandidateIndex >= Math.max((response.candidate_rows || state.candidateRows).length, 1)) {
      state.selectedCandidateIndex = Math.max((response.candidate_rows || state.candidateRows).length - 1, 0);
    }
    applyTaskRefresh(response);
  });
  document.getElementById("candidatePrevBtn").addEventListener("click", () => {
    state.selectedCandidateIndex = Math.max(0, state.selectedCandidateIndex - 1);
    renderCandidateList();
    renderCandidateDetail();
  });
  document.getElementById("candidateNextBtn").addEventListener("click", () => {
    state.selectedCandidateIndex = Math.min(state.candidateRows.length - 1, state.selectedCandidateIndex + 1);
    renderCandidateList();
    renderCandidateDetail();
  });
}

// ── 層級標籤對照 ─────────────────────────────────────────────
const SUBSIDIARY_LABELS = {
  0: "集團本級", 1: "一級子公司", 2: "二級子公司",
  3: "三級子公司", 4: "四級子公司",
};

// ── 拖曳重新掛父層 ────────────────────────────────────────────
let _dragNodeId = null;
let _dragNodeIds = [];

function isAncestor(candidateAncestorId, nodeId) {
  // 判斷 candidateAncestorId 是否為 nodeId 的祖先（防止循環）
  const byId = {};
  state.masterRows.forEach((r) => { byId[r.node_id] = r; });
  let cur = byId[nodeId];
  while (cur && cur.chart1_parent) {
    if (cur.chart1_parent === candidateAncestorId) return true;
    cur = byId[cur.chart1_parent];
  }
  return false;
}

function getVisibleResultRows() {
  return flattenTree(buildTree(state.masterRows)).filter((row) => (Number(row.chart1_level) || 0) > 0);
}

function topLevelDragIds(nodeIds) {
  const selected = new Set(nodeIds);
  return nodeIds.filter((nodeId) => {
    const row = state.masterRows.find((r) => r.node_id === nodeId);
    let parentId = row?.chart1_parent || "";
    while (parentId) {
      if (selected.has(parentId)) return false;
      const parent = state.masterRows.find((r) => r.node_id === parentId);
      parentId = parent?.chart1_parent || "";
    }
    return true;
  });
}

function collectDescendantIds(nodeId) {
  const ids = [];
  function walk(id) {
    state.masterRows
      .filter((row) => row.chart1_parent === id)
      .forEach((child) => {
        ids.push(child.node_id);
        walk(child.node_id);
      });
  }
  walk(nodeId);
  return ids;
}

function invalidBulkDrop(targetId, dragIds) {
  return dragIds.some((dragId) => dragId === targetId || isAncestor(dragId, targetId));
}

function currentDragIds(rowId) {
  const visibleIds = getVisibleResultRows().map((row) => row.node_id);
  if (state.selectedResultNodeIds.has(rowId)) {
    const selectedVisibleIds = visibleIds.filter((nodeId) => state.selectedResultNodeIds.has(nodeId));
    return topLevelDragIds(selectedVisibleIds.length ? selectedVisibleIds : [rowId]);
  }
  return [rowId];
}

async function saveWholeTaskState() {
  const payload = {
    master_rows: JSON.parse(JSON.stringify(state.masterRows)),
    review_rows: JSON.parse(JSON.stringify(state.reviewRows)),
    candidate_rows: JSON.parse(JSON.stringify(state.candidateRows)),
    review_decisions: JSON.parse(JSON.stringify(state.reviewDecisions)),
    candidate_decisions: JSON.parse(JSON.stringify(state.candidateDecisions)),
  };
  const result = await apiPost(`/api/tasks/${state.taskId}/replace-state`, payload);
  applyTaskRefresh(result);
}

async function reparentNode(nodeId, newParentId) {
  const before = snapshotTaskState();
  const byId = {};
  state.masterRows.forEach((r) => { byId[r.node_id] = r; });
  const node = byId[nodeId];
  const newParent = byId[newParentId];
  if (!node || !newParent) return;

  const oldLevel = Number(node.chart1_level) || 0;
  const newLevel = (Number(newParent.chart1_level) || 0) + 1;
  const diff = newLevel - oldLevel;

  // 遞迴更新節點及所有後代的層級
  function cascadeLevel(id) {
    const r = byId[id];
    if (!r) return;
    const lv = (Number(r.chart1_level) || 0) + diff;
    r.chart1_level = lv;
    r.subsidiary_level_label = SUBSIDIARY_LABELS[lv] || `${lv}級子公司`;
    state.masterRows
      .filter((c) => c.chart1_parent === id)
      .forEach((c) => cascadeLevel(c.node_id));
  }

  node.chart1_parent = newParentId;
  node.chart1_parent_name = newParent.canonical_name || newParent.chart1_name || "";
  cascadeLevel(nodeId);

  renderResults();

  // 蒐集所有受影響節點（dragged + descendants）一起存檔
  const changed = [];
  function collect(id) {
    const r = byId[id];
    if (!r) return;
    changed.push(r);
    state.masterRows.filter((c) => c.chart1_parent === id).forEach((c) => collect(c.node_id));
  }
  collect(nodeId);

  await Promise.all(changed.map((r) =>
    apiPost(`/api/tasks/${state.taskId}/update-row`, {
      node_id: r.node_id,
      chart1_parent:       r === node ? newParentId : r.chart1_parent,
      chart1_parent_name:  r === node ? node.chart1_parent_name : r.chart1_parent_name,
      chart1_level:        String(r.chart1_level),
      subsidiary_level_label: r.subsidiary_level_label,
    }).catch((err) => console.error("reparent save failed", err))
  ));
  pushUndoSnapshot(before);
}

async function reparentNodes(nodeIds, newParentId) {
  const dragIds = topLevelDragIds(nodeIds);
  if (dragIds.length <= 1) return reparentNode(dragIds[0], newParentId);
  if (invalidBulkDrop(newParentId, dragIds)) return;

  const before = snapshotTaskState();
  const byId = {};
  state.masterRows.forEach((r) => { byId[r.node_id] = r; });
  const newParent = byId[newParentId];
  if (!newParent) return;

  dragIds.forEach((nodeId) => {
    const node = byId[nodeId];
    if (!node) return;
    const oldLevel = Number(node.chart1_level) || 0;
    const newLevel = (Number(newParent.chart1_level) || 0) + 1;
    const diff = newLevel - oldLevel;

    node.chart1_parent = newParentId;
    node.chart1_parent_name = newParent.canonical_name || newParent.chart1_name || "";

    function cascadeLevel(id) {
      const r = byId[id];
      if (!r) return;
      const lv = (Number(r.chart1_level) || 0) + diff;
      r.chart1_level = lv;
      r.subsidiary_level_label = SUBSIDIARY_LABELS[lv] || `${lv}級子公司`;
      state.masterRows
        .filter((c) => c.chart1_parent === id)
        .forEach((c) => cascadeLevel(c.node_id));
    }
    cascadeLevel(nodeId);
  });

  normalizeSiblingOrder();
  await saveWholeTaskState();
  pushUndoSnapshot(before);
}

function normalizeSiblingOrder() {
  const groups = new Map();
  state.masterRows.forEach((row) => {
    const key = row.chart1_parent || "__root__";
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(row);
  });
  groups.forEach((rows) => {
    rows
      .sort((a, b) => Number(a.sort_index || 0) - Number(b.sort_index || 0))
      .forEach((row, idx) => { row.sort_index = idx + 1; });
  });
}

async function reorderSiblingNode(nodeId, targetId, zone) {
  const before = snapshotTaskState();
  const byId = {};
  state.masterRows.forEach((r) => { byId[r.node_id] = r; });
  const node = byId[nodeId];
  const target = byId[targetId];
  if (!node || !target) return;

  const parentId = target.chart1_parent || "";
  node.chart1_parent = parentId;
  node.chart1_parent_name = target.chart1_parent_name || "";
  node.chart1_level = target.chart1_level;
  node.subsidiary_level_label = target.subsidiary_level_label;

  const siblings = state.masterRows
    .filter((row) => (row.chart1_parent || "") === parentId && row.node_id !== nodeId)
    .sort((a, b) => Number(a.sort_index || 0) - Number(b.sort_index || 0));
  const targetIndex = siblings.findIndex((row) => row.node_id === targetId);
  const insertIndex = zone === "top" ? targetIndex : targetIndex + 1;
  siblings.splice(Math.max(insertIndex, 0), 0, node);
  siblings.forEach((row, idx) => { row.sort_index = idx + 1; });
  normalizeSiblingOrder();
  renderResults();

  await saveWholeTaskState();
  pushUndoSnapshot(before);
}

async function reorderSiblingNodes(nodeIds, targetId, zone) {
  const dragIds = topLevelDragIds(nodeIds);
  if (dragIds.length <= 1) return reorderSiblingNode(dragIds[0], targetId, zone);
  if (invalidBulkDrop(targetId, dragIds)) return;

  const before = snapshotTaskState();
  const byId = {};
  state.masterRows.forEach((r) => { byId[r.node_id] = r; });
  const target = byId[targetId];
  if (!target) return;
  const parentId = target.chart1_parent || "";
  const parentName = target.chart1_parent_name || "";
  const targetLevel = target.chart1_level;
  const targetLabel = target.subsidiary_level_label;
  const moving = getVisibleResultRows()
    .filter((row) => dragIds.includes(row.node_id))
    .map((row) => byId[row.node_id])
    .filter(Boolean);

  moving.forEach((node) => {
    const oldLevel = Number(node.chart1_level) || 0;
    const newLevel = Number(targetLevel) || 0;
    const diff = newLevel - oldLevel;
    node.chart1_parent = parentId;
    node.chart1_parent_name = parentName;
    node.chart1_level = targetLevel;
    node.subsidiary_level_label = targetLabel;
    collectDescendantIds(node.node_id).forEach((descId) => {
      const desc = byId[descId];
      if (!desc) return;
      const lv = (Number(desc.chart1_level) || 0) + diff;
      desc.chart1_level = lv;
      desc.subsidiary_level_label = SUBSIDIARY_LABELS[lv] || `${lv}級子公司`;
    });
  });

  const movingSet = new Set(moving.map((row) => row.node_id));
  const siblings = state.masterRows
    .filter((row) => (row.chart1_parent || "") === parentId && !movingSet.has(row.node_id))
    .sort((a, b) => Number(a.sort_index || 0) - Number(b.sort_index || 0));
  const targetIndex = siblings.findIndex((row) => row.node_id === targetId);
  const insertIndex = zone === "top" ? targetIndex : targetIndex + 1;
  siblings.splice(Math.max(insertIndex, 0), 0, ...moving);
  siblings.forEach((row, idx) => { row.sort_index = idx + 1; });
  normalizeSiblingOrder();
  await saveWholeTaskState();
  pushUndoSnapshot(before);
}

// ── 樹狀結構建立 ─────────────────────────────────────────────
function buildTree(rows) {
  const byId = {};
  rows.forEach((r) => { byId[r.node_id] = { ...r, children: [] }; });
  const roots = [];
  rows.forEach((r) => {
    const parent = r.chart1_parent && byId[r.chart1_parent];
    if (parent) parent.children.push(byId[r.node_id]);
    else roots.push(byId[r.node_id]);
  });
  Object.values(byId).forEach((node) => {
    node.children.sort((a, b) => Number(a.sort_index || 0) - Number(b.sort_index || 0));
  });
  roots.sort((a, b) => Number(a.sort_index || 0) - Number(b.sort_index || 0));
  return roots;
}

function flattenTree(nodes, depth = 0, result = []) {
  nodes.forEach((node) => {
    result.push({ ...node, _depth: depth });
    if (node.children?.length) flattenTree(node.children, depth + 1, result);
  });
  return result;
}

function renderAddCompanyParentOptions() {
  if (!elements.addCompanyParent) return;
  const rows = flattenTree(buildTree(state.masterRows));
  const root = getRootRow();
  const options = [];
  if (root) {
    options.push(`<option value="${root.node_id}">掛在${getGroupName()}底下（一級子公司）</option>`);
  } else {
    options.push(`<option value="">設為頂層公司</option>`);
  }
  rows
    .filter((row) => (Number(row.chart1_level) || 0) > 0)
    .forEach((row) => {
      const level = Number(row.chart1_level) || 0;
      const prefix = "　".repeat(Math.max(level - 1, 0));
      const label = row.canonical_name || row.chart1_name || "未命名公司";
      options.push(`<option value="${row.node_id}">${prefix}${label}</option>`);
    });
  elements.addCompanyParent.innerHTML = options.join("");
}

function setAddCompanyPanel(open) {
  if (!elements.addCompanyPanel) return;
  elements.addCompanyPanel.classList.toggle("active", open);
  if (open && elements.bulkAddPanel) elements.bulkAddPanel.classList.remove("active");
  if (open) {
    renderAddCompanyParentOptions();
    elements.addCompanyName.value = "";
    elements.addCompanyShare.value = "";
    elements.addCompanyName.focus();
  }
}

function setBulkAddPanel(open) {
  if (!elements.bulkAddPanel) return;
  elements.bulkAddPanel.classList.toggle("active", open);
  if (open && elements.addCompanyPanel) elements.addCompanyPanel.classList.remove("active");
  if (open) {
    elements.bulkAddInput.value = "";
    elements.bulkAddInput.focus();
  }
}

async function addCompanyToResults() {
  if (!state.taskId) return;
  const name = elements.addCompanyName?.value.trim() || "";
  if (!name) {
    elements.addCompanyName?.focus();
    return;
  }
  const originalText = elements.saveAddCompanyBtn.textContent;
  elements.saveAddCompanyBtn.disabled = true;
  elements.saveAddCompanyBtn.textContent = "加入中...";
  try {
    const before = snapshotTaskState();
    const result = await apiPost(`/api/tasks/${state.taskId}/add-row`, {
      canonical_name: name,
      chart1_parent: elements.addCompanyParent?.value || "",
      actual_controller_share: elements.addCompanyShare?.value.trim() || "",
    });
    applyTaskRefresh(result);
    pushUndoSnapshot(before);
    setAddCompanyPanel(false);
    setView("results");
    document.querySelector(`tr[data-node-id="${result.added_node_id}"]`)?.scrollIntoView({ behavior: "smooth", block: "center" });
  } catch (error) {
    console.error("新增公司失敗", error);
  } finally {
    elements.saveAddCompanyBtn.disabled = false;
    elements.saveAddCompanyBtn.textContent = originalText;
  }
}

function parseBulkAddLines(text) {
  const lines = String(text || "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
  return lines.map((line, index) => {
    const parts = line.split(/[，,]/).map((part) => part.trim());
    if (!parts[0]) {
      throw new Error(`第 ${index + 1} 行缺少公司名稱`);
    }
    return {
      canonical_name: parts[0],
      parent_name: parts[1] || "",
      actual_controller_share: parts[2] || "",
    };
  });
}

async function bulkAddCompaniesToResults() {
  if (!state.taskId) return;
  const items = parseBulkAddLines(elements.bulkAddInput?.value || "");
  if (!items.length) {
    elements.bulkAddInput?.focus();
    throw new Error("請先輸入至少一行公司資料");
  }
  const originalText = elements.saveBulkAddBtn.textContent;
  elements.saveBulkAddBtn.disabled = true;
  elements.saveBulkAddBtn.textContent = "加入中...";
  try {
    const before = snapshotTaskState();
    const result = await apiPost(`/api/tasks/${state.taskId}/bulk-add-rows`, { items });
    applyTaskRefresh(result);
    pushUndoSnapshot(before);
    setBulkAddPanel(false);
    setView("results");
  } finally {
    elements.saveBulkAddBtn.disabled = false;
    elements.saveBulkAddBtn.textContent = originalText;
  }
}

// ── 資本額格式化 ───────────────────────────────────────────────
function formatCapital(str) {
  if (!str || str === "—") return str;
  return str.replace(/^(\d+)/, (_, n) => parseInt(n, 10).toLocaleString("en-US"));
}

// ── 公司名稱 inline 編輯（contenteditable）────────────────────
function attachNameEdit(td, row) {
  td.addEventListener("click", (e) => {
    const t = e.target;
    if (t.classList.contains("drag-handle")) return;
    const nameSpan = td.querySelector(".company-name");
    if (!nameSpan || nameSpan.contentEditable === "true") return;
    const curName = nameSpan.textContent;
    nameSpan.contentEditable = "true";
    nameSpan.spellcheck = false;
    nameSpan.focus();
    const sel = window.getSelection();
    const range = document.createRange();
    range.selectNodeContents(nameSpan);
    range.collapse(false);
    sel.removeAllRanges();
    sel.addRange(range);
    let finished = false;
    const finish = async (save) => {
      if (finished) return;
      finished = true;
      nameSpan.contentEditable = "false";
      const newVal = nameSpan.textContent.trim();
      if (save && newVal && newVal !== curName) {
        const before = snapshotTaskState();
        row.canonical_name = newVal;
        try {
          const result = await apiPost(`/api/tasks/${state.taskId}/update-row`, {
            node_id: row.node_id, canonical_name: newVal,
          });
          applyTaskRefresh(result);
          pushUndoSnapshot(before);
        } catch (err) {
          console.error(err);
          nameSpan.textContent = curName;
          row.canonical_name = curName;
        }
      } else if (!save) {
        nameSpan.textContent = curName;
      }
    };
    nameSpan.addEventListener("blur", () => finish(true), { once: true });
    nameSpan.addEventListener("keydown", (ev) => {
      if (ev.key === "Enter")  { ev.preventDefault(); finish(true); }
      if (ev.key === "Escape") { finish(false); }
    }, { once: true });
  });
}

function unformatCapital(str) {
  return (str || "").replace(/,/g, "");
}

// 連動更新欄位（修改一個，同值的全部跟著改）
const CASCADE_FIELDS = new Set(["legal_representative"]);
const BLANK_DISPLAY_FIELDS = new Set(["role_label", "chart_note"]);

// ── 行內編輯 ──────────────────────────────────────────────────
function makeEditable(cell, row, field, displayValue) {
  cell.title = "點擊編輯";
  cell.addEventListener("click", () => {
    if (cell.querySelector("input")) return;

    // 輸入框顯示的是「原始值」（無格式）
    const rawOriginal = field === "registered_capital"
      ? unformatCapital(row[field] || "")
      : (row[field] || "");

    cell.innerHTML = `<input class="cell-input" value="${rawOriginal}" />`;
    const input = cell.querySelector("input");
    input.focus();
    input.select();

    const save = async () => {
      const newRaw = input.value.trim();

      // 顯示格式化後的值
      const display = field === "registered_capital" ? formatCapital(newRaw) : newRaw;
      cell.textContent = display || (BLANK_DISPLAY_FIELDS.has(field) ? "" : "—");

      if (newRaw === rawOriginal) return; // 沒有改變

      try {
        const before = snapshotTaskState();
        let result;
        if (CASCADE_FIELDS.has(field) && rawOriginal) {
          // 連動：同名全部更新
          result = await apiPost(`/api/tasks/${state.taskId}/update-row`, {
            cascade: true,
            field,
            original_value: rawOriginal,
            new_value: newRaw,
          });
        } else {
          result = await apiPost(`/api/tasks/${state.taskId}/update-row`, {
            node_id: row.node_id,
            [field]: newRaw,
          });
        }
        applyTaskRefresh(result);
        pushUndoSnapshot(before);
        // 連動時重新渲染整張表
        if (CASCADE_FIELDS.has(field)) renderResults();
      } catch (e) {
        console.error("儲存失敗", e);
        cell.textContent = displayValue || (BLANK_DISPLAY_FIELDS.has(field) ? "" : "—");
      }
    };

    input.addEventListener("blur", () => {
      const display = field === "registered_capital" ? formatCapital(rawOriginal) : rawOriginal;
      cell.textContent = display || (BLANK_DISPLAY_FIELDS.has(field) ? "" : "—");
    });
    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") { e.preventDefault(); save(); }
      if (e.key === "Escape") { cell.textContent = displayValue || "—"; }
    });
  });
}

function renderResults() {
  const rows = getVisibleResultRows();
  const companyCount = Math.max(state.masterRows.length - (getRootRow() ? 1 : 0), 0);

  elements.resultTableTitle.textContent = `${getGroupName()}共 ${companyCount} 間公司`;
  if (elements.addCompanyBtn) elements.addCompanyBtn.disabled = !state.taskId;
  if (elements.bulkAddBtn) elements.bulkAddBtn.disabled = !state.taskId;
  renderAddCompanyParentOptions();

  // ── 動態層級欄數 ────────────────────────────────────────────
  const maxLevel = Math.max(0, ...state.masterRows.map((r) => Number(r.chart1_level) || 0));
  const LEVEL_HEADERS = { 1: "一級子公司", 2: "二級子公司", 3: "三級子公司", 4: "四級子公司" };

  // 更新表頭
  const theadTr = document.querySelector("#results table thead tr");
  let headHtml = `<th class="del-col"><input id="resultSelectAll" type="checkbox" /></th><th class="del-col"></th>`;
  for (let lv = 1; lv <= maxLevel; lv++) {
    headHtml += `<th class="level-col" title="公司名稱可點擊修改；層級關係用拖曳調整">${LEVEL_HEADERS[lv] || `${lv}級子公司`} ✏️</th>`;
  }
  headHtml += `<th class="editable-col">法定代表人 ✏️</th>
    <th class="editable-col">資本額 ✏️</th>
    <th class="editable-col">成立日期 ✏️</th>
    <th class="editable-col">持股% ✏️</th>
    <th class="editable-col">定位 ✏️</th>
    <th class="status-col">備註 ✏️</th>`;
  theadTr.innerHTML = headHtml;
  elements.resultSelectAll = document.getElementById("resultSelectAll");
  updateSelectedRowsHint(rows.filter((row) => state.selectedResultNodeIds.has(row.node_id)).length);
  elements.resultSelectAll?.addEventListener("change", (event) => {
    if (event.target.checked) {
      rows.forEach((row) => state.selectedResultNodeIds.add(row.node_id));
    } else {
      rows.forEach((row) => state.selectedResultNodeIds.delete(row.node_id));
    }
    renderResults();
  });

  // ── 分組底色（依一級祖先交替）──────────────────────────────
  const byId = {};
  state.masterRows.forEach((r) => { byId[r.node_id] = r; });
  const GROUP_COLORS = ["#f0f9ff", "#f0fdf4", "#fefce8", "#fdf4ff", "#fff7ed"];
  const groupBgMap = {};
  let ci = 0;
  function assignBg(nodeId, color) {
    groupBgMap[nodeId] = color;
    state.masterRows.filter((r) => r.chart1_parent === nodeId).forEach((c) => assignBg(c.node_id, color));
  }
  state.masterRows.filter((r) => (Number(r.chart1_level) || 0) === 0).forEach((r) => assignBg(r.node_id, "#ffffff"));
  state.masterRows.filter((r) => (Number(r.chart1_level) || 0) === 1).forEach((r) => assignBg(r.node_id, GROUP_COLORS[ci++ % GROUP_COLORS.length]));

  // ── 渲染每一行 ──────────────────────────────────────────────
  elements.resultTableBody.innerHTML = "";
  rows.forEach((row) => {
    const level = Number(row.chart1_level) || 0;
    const statusClass = row.node_status === "manual_added" ? "status-manual"
      : row.node_status === "enriched" ? "status-enriched"
      : row.node_status === "review_match" ? "status-review" : "status-slate";

    const tr = document.createElement("tr");
    tr.className = statusClass;
    tr.dataset.nodeId = row.node_id;
    const bg = groupBgMap[row.node_id];
    if (bg) tr.style.backgroundColor = bg;

    // 刪除按鈕
    const selectTd = document.createElement("td");
    selectTd.className = "del-td";
    const selectInput = document.createElement("input");
    selectInput.type = "checkbox";
    selectInput.checked = state.selectedResultNodeIds.has(row.node_id);
    selectInput.addEventListener("change", () => {
      if (selectInput.checked) state.selectedResultNodeIds.add(row.node_id);
      else state.selectedResultNodeIds.delete(row.node_id);
      updateResultSelectAllState(rows);
    });
    selectTd.appendChild(selectInput);
    tr.appendChild(selectTd);

    // 刪除按鈕
    const delTd = document.createElement("td");
    delTd.className = "del-td";
    const delBtn = document.createElement("button");
    delBtn.className = "row-del-btn";
    delBtn.title = "從主表移除此公司";
    delBtn.textContent = "✕";
    delBtn.addEventListener("click", async (e) => {
      e.stopPropagation();
      if (!confirm(`確定要從主表移除「${row.canonical_name || row.chart1_name}」嗎？`)) return;
      try {
        const before = snapshotTaskState();
        const result = await apiPost(`/api/tasks/${state.taskId}/delete-row`, { node_id: row.node_id });
        applyTaskRefresh(result);
        pushUndoSnapshot(before);
      } catch (err) { console.error("刪除失敗", err); }
    });
    delTd.appendChild(delBtn);
    tr.appendChild(delTd);

    // ── 拖曳事件（整行） ────────────────────────────────────
    tr.draggable = true;
    let dropZone = "middle";
    tr.addEventListener("dragstart", (e) => {
      _dragNodeIds = currentDragIds(row.node_id);
      _dragNodeId = _dragNodeIds[0] || row.node_id;
      e.dataTransfer.effectAllowed = "move";
      e.dataTransfer.setData("text/plain", _dragNodeId);
      e.dataTransfer.setData("application/json", JSON.stringify(_dragNodeIds));
      const ghost = document.createElement("div");
      ghost.textContent = _dragNodeIds.length > 1
        ? `移動 ${_dragNodeIds.length} 間公司`
        : (row.canonical_name || row.chart1_name || "");
      ghost.style.cssText = "position:fixed;top:-200px;left:0;background:#4f46e5;color:#fff;padding:5px 14px;border-radius:99px;font-size:13px;font-weight:600;white-space:nowrap;";
      document.body.appendChild(ghost);
      e.dataTransfer.setDragImage(ghost, ghost.offsetWidth / 2, 18);
      setTimeout(() => { ghost.remove(); tr.classList.add("row-dragging"); }, 0);
    });
    tr.addEventListener("dragend", () => {
      tr.classList.remove("row-dragging");
      document.querySelectorAll(".drag-target").forEach((el) => el.classList.remove("drag-target"));
      document.getElementById("drag-tooltip")?.remove();
      _dragNodeId = null;
      _dragNodeIds = [];
    });
    tr.addEventListener("dragover", (e) => {
      if (!_dragNodeId || invalidBulkDrop(row.node_id, _dragNodeIds.length ? _dragNodeIds : [_dragNodeId])) return;
      e.preventDefault();
      e.dataTransfer.dropEffect = "move";
      document.querySelectorAll(".drag-target").forEach((el) => el.classList.remove("drag-target"));
      tr.classList.remove("drag-insert-top", "drag-insert-bottom");
      tr.classList.add("drag-target");
      const rect = tr.getBoundingClientRect();
      const ratio = (e.clientY - rect.top) / Math.max(rect.height, 1);
      if (ratio < 0.25) dropZone = "top";
      else if (ratio > 0.75) dropZone = "bottom";
      else dropZone = "middle";
      if (dropZone === "top") tr.classList.add("drag-insert-top");
      if (dropZone === "bottom") tr.classList.add("drag-insert-bottom");
      let tip = document.getElementById("drag-tooltip");
      if (!tip) { tip = document.createElement("div"); tip.id = "drag-tooltip"; document.body.appendChild(tip); }
      tip.textContent = dropZone === "middle"
        ? `↳ 移到「${row.canonical_name || row.chart1_name}」的下一個層級${_dragNodeIds.length > 1 ? `（${_dragNodeIds.length} 間）` : ""}`
        : `↕ 插入為「${row.canonical_name || row.chart1_name}」${dropZone === "top" ? "上方同層級" : "下方同層級"}${_dragNodeIds.length > 1 ? `（${_dragNodeIds.length} 間）` : ""}`;
      tip.style.left = e.clientX + "px";
      tip.style.top  = e.clientY + "px";
    });
    tr.addEventListener("dragleave", (e) => {
      if (!tr.contains(e.relatedTarget)) {
        tr.classList.remove("drag-target", "drag-insert-top", "drag-insert-bottom");
      }
    });
    tr.addEventListener("drop", async (e) => {
      e.preventDefault();
      tr.classList.remove("drag-target", "drag-insert-top", "drag-insert-bottom");
      document.getElementById("drag-tooltip")?.remove();
      const fallbackId = e.dataTransfer.getData("text/plain");
      let draggedIds = _dragNodeIds.length ? _dragNodeIds : [fallbackId];
      try {
        const parsed = JSON.parse(e.dataTransfer.getData("application/json") || "[]");
        if (Array.isArray(parsed) && parsed.length) draggedIds = parsed;
      } catch {}
      draggedIds = topLevelDragIds(draggedIds.filter(Boolean));
      if (!draggedIds.length || invalidBulkDrop(row.node_id, draggedIds)) return;
      if (dropZone === "middle") {
        await reparentNodes(draggedIds, row.node_id);
      } else {
        await reorderSiblingNodes(draggedIds, row.node_id, dropZone);
      }
      document.querySelector(`tr[data-node-id="${draggedIds[0]}"]`)?.scrollIntoView({ behavior: "smooth", block: "center" });
    });

    // ── 層級欄：公司名在對應欄，其餘空白 ────────────────────
    const curName = row.canonical_name || row.chart1_name || "";
    for (let lv = 1; lv <= maxLevel; lv++) {
      const td = document.createElement("td");
      if (lv === level) {
        td.className = "tree-name-cell";
        td.innerHTML = `<span class="drag-handle" title="拖曳調整層級">⠿</span><span class="company-name editable-name" title="點擊編輯名稱">${curName}</span>`;
        attachNameEdit(td, row);
      } else {
        td.className = "level-empty-td";
        td.title = "此公司不在這個層級；要調整層級請拖曳公司名稱列";
      }
      tr.appendChild(td);
    }

    // 資料欄
    [
      { key: "legal_representative",    display: row.legal_representative },
      { key: "registered_capital",      display: formatCapital(row.registered_capital) },
      { key: "established_date",        display: row.established_date },
      { key: "actual_controller_share", display: row.actual_controller_share },
      { key: "role_label",              display: row.role_label },
      { key: "chart_note",              display: row.chart_note },
    ].forEach(({ key, display }) => {
      const td = document.createElement("td");
      td.textContent = display || (BLANK_DISPLAY_FIELDS.has(key) ? "" : "—");
      td.className = "editable-cell";
      makeEditable(td, row, key, display);
      tr.appendChild(td);
    });

    elements.resultTableBody.appendChild(tr);
  });
  updateResultSelectAllState(rows);
}

// 輪詢任務狀態，回呼 onStatus(task) 讓呼叫端更新 UI
// 終止條件：ready / error / chart2_error
async function pollTask(taskId, onStatus) {
  const MAX_MS = 20 * 60 * 1000;
  const INTERVAL = 4000;
  const start = Date.now();
  while (Date.now() - start < MAX_MS) {
    await new Promise((r) => setTimeout(r, INTERVAL));
    const task = await apiGet(`/api/tasks/${taskId}`);
    if (onStatus) onStatus(task);
    if (task.status === "ready") return task;
    if (task.status === "chart2_error") return task;
    if (task.status === "chart2_ocr_done") return task;  // 圖二 OCR 完成，等用戶確認
    if (task.status === "cancelled" || task.status === "cancel_requested") return task;
    if (task.status === "error") throw new Error(task.error || "AI 辨識失敗，請重試");
    // processing / chart1_ready / processing_chart2 → 繼續等
  }
  throw new Error("分析逾時（超過 20 分鐘），請重試或裁切圖片後再上傳");
}

async function cancelCurrentTask() {
  if (!state.taskId) return;
  await apiPost(`/api/tasks/${state.taskId}/cancel`, {});
}

// ── 工作區渲染 ────────────────────────────────────────────────
// phase: "idle"|"uploading"|"processing"|"chart1_ready"|"processing_chart2"
//        |"chart2_confirm"|"ready"|"chart2_error"|"cancelled"|"error"
function renderWorkspace(phase, opts = {}) {
  const el = document.getElementById("workspaceContent");
  if (!el) return;
  trackWorkspaceActivity(phase, opts);

  // ── 特殊：圖二確認畫面 ───────────────────────────────────
  if (phase === "chart2_confirm") {
    const task = opts.task || {};
    const chart2Raw = task.chart2_raw || [];
    const chart2Progress = task.chart2_progress || {};
    const chart2Quality = task.analysis_diagnostics?.chart2_quality || {};
    const chart1Count = (task.master_rows || []).length;
    const chart2Count = chart2Raw.length;
    const failedChunks = Array.isArray(chart2Progress.failed_chunks) ? chart2Progress.failed_chunks.length : 0;
    const discrepancy = chart1Count > 0 && chart2Count < chart1Count * 0.5;
    const qualityNotes = Array.isArray(chart2Quality.notes) ? chart2Quality.notes : [];
    const qualityRisk = chart2Quality.review_required || chart2Quality.level === "low";
    const success = evaluateRecognitionSuccess({ masterRows: task.master_rows || [], chart2Rows: chart2Raw });

    const companyListHtml = chart2Raw.length
      ? chart2Raw.map((c) => `<li class="ws-company-item">${c.company || "（未知）"}</li>`).join("")
      : `<li class="ws-company-item ws-company-empty">（無辨識結果）</li>`;

    el.innerHTML = `
      <p class="workspace-eyebrow">需要確認</p>
      <div class="ws-confirm-alert">
        <strong>圖二辨識已完成，請確認名單後開始配對</strong>
        <span>${failedChunks ? `已保留可用結果；有 ${failedChunks} 個分塊需要人工確認。` : "這不是失敗；系統已先停下來，等你確認圖二公司清單是否可以繼續。"}</span>
      </div>
      <div class="ws-confirm-counts">
        <div class="ws-cc-stat">
          <span class="ws-cc-label">圖一結構</span>
          <span class="ws-cc-val">${chart1Count} 家</span>
        </div>
        <div class="ws-cc-arrow">→</div>
        <div class="ws-cc-stat">
          <span class="ws-cc-label">圖二辨識</span>
          <span class="ws-cc-val ${discrepancy ? "ws-cc-warn" : ""}">${chart2Count} 家</span>
        </div>
      </div>
      ${recognitionSuccessHtml(success)}
      ${qualityRisk ? `<div class="ws-warn-box">圖二辨識品質偏低：${qualityNotes.join("；") || "建議重新上傳更清晰或裁切後的圖二。"}</div>` : ""}
      ${discrepancy ? `<div class="ws-warn-box">圖二辨識到的公司數量明顯少於圖一，建議確認圖片清晰度後再繼續。</div>` : ""}
      ${failedChunks ? `<div class="ws-warn-box">圖二已完成可用辨識結果；少數分塊未完整抽出細節，仍可先進入配對並在主表補正。</div>` : ""}
      <p class="ws-company-list-label">圖二辨識到的公司（${chart2Count} 筆）</p>
      <ul class="ws-company-list">${companyListHtml}</ul>
      <div class="ws-confirm-actions">
        <button class="ws-confirm-btn" id="wsConfirmChart2Btn">確認圖二，開始配對</button>
        <label class="ws-reupload-label">
          重新上傳圖二
          <input type="file" id="wsChart2ReuploadInput" accept="image/*" style="display:none" />
        </label>
      </div>`;

    document.getElementById("wsConfirmChart2Btn")?.addEventListener("click", () => {
      confirmChart2Match(task.id, task);
    });

    document.getElementById("wsChart2ReuploadInput")?.addEventListener("change", async (e) => {
      const file = e.target.files?.[0];
      if (!file || !task.id) return;
      renderWorkspace("processing_chart2", { msg: "重新上傳圖二中…" });
      try {
        const fd = new FormData();
        fd.append("chart2", file);
        await apiPost(`/api/tasks/${task.id}/analyze-chart2`, fd, true);
        renderWorkspace("processing_chart2", { msg: "圖二辨識中…" });
        const updated = await pollTask(task.id, (t) => {
          renderWorkspace("processing_chart2", { progress: t.chart2_progress, summary: t.summary });
        });
        if (updated.status === "chart2_ocr_done") {
          renderWorkspace("chart2_confirm", { task: updated });
        } else if (updated.status === "chart2_error") {
          renderWorkspace("chart2_error", { taskId: updated.id, error: updated.error, summary: updated.summary, diagnostics: updated.analysis_diagnostics });
        } else if (updated.status === "cancelled" || updated.status === "cancel_requested") {
          renderWorkspace("cancelled", { summary: updated.summary });
        } else {
          hydrateTask(updated);
          renderWorkspace("ready", { summary: updated.summary });
        }
      } catch (err) {
        renderWorkspace("error", { error: err.message });
      }
    });
    return;
  }

  const c1name = state.chart1File?.name || "—";
  const c2name = state.chart2File?.name || "—";

  // ── 每個步驟的狀態 ────────────────────────────────────────
  function ss(key) {
    switch (phase) {
      case "idle":              return "pending";
      case "uploading":         return key === "upload" ? "active" : "pending";
      case "processing":        return key === "upload" ? "done" : key === "chart1" ? "active" : "pending";
      case "chart1_ready":      return key === "enrich" ? "active" : "done";
      case "processing_chart2": return key === "enrich" ? "active" : "done";
      case "chart2_confirm":    return key === "enrich" ? "active" : "done";
      case "ready":             return "done";
      case "chart2_error":      return key === "enrich" ? "error" : "done";
      case "cancelled":         return key === "enrich" ? "error" : "done";
      case "error":             return key === "upload" ? "error" : "pending";
      default:                  return "pending";
    }
  }

  function detail(key) {
    if (key === "upload") {
      if (phase === "idle") return "等待開始";
      if (phase === "uploading") return "上傳圖一與圖二至伺服器…";
      return `${c1name} · ${c2name}`;
    }
    if (key === "chart1") {
      if (phase === "idle" || phase === "uploading") return "等待上傳完成";
      if (phase === "processing") return opts.msg || "辨識中…";
      if (phase === "chart1_ready" || phase === "processing_chart2" || phase === "chart2_confirm" || phase === "ready") {
        const s = opts.summary || {};
        const quality = s.chart1_quality || {};
        const rescueNote = quality.rescue_used ? "，已補強公司名稱" : "";
        return `完成 — ${s.master_count ?? "?"} 間公司${rescueNote}`;
      }
      if (phase === "chart2_error") {
        const s = opts.summary || {};
        const quality = s.chart1_quality || {};
        const rescueNote = quality.rescue_used ? "，已補強公司名稱" : "";
        return `完成 — ${s.master_count ?? "?"} 間公司${rescueNote}`;
      }
      return "—";
    }
    if (key === "enrich") {
      if (phase === "idle" || phase === "uploading" || phase === "processing") return "等待圖一完成";
      if (phase === "chart1_ready") return "準備辨識圖二…";
      if (phase === "processing_chart2") return opts.progress ? chart2ProgressText(opts.progress) : (opts.msg || "辨識中…");
      if (phase === "chart2_confirm") return "圖二已完成，等待確認";
      if (phase === "ready") return "補充完成";
      if (phase === "chart2_error") return "辨識失敗，可重新上傳圖二";
      if (phase === "cancelled") return "任務已取消";
      return "—";
    }
    return "";
  }

  const icons = { pending: "·", active: "…", done: "✓", error: "!" };
  const stepDefs = [
    { key: "upload", title: "上傳圖片" },
    { key: "chart1", title: "圖一辨識（結構骨架）" },
    { key: "enrich", title: "圖二辨識（補充資訊）" },
  ];

  const stepsHtml = stepDefs.map((s) => {
    const st = ss(s.key);
    const d = detail(s.key);
    return `
      <li class="ws-step ${st}">
        <div class="ws-step-icon">${icons[st]}</div>
        <div class="ws-step-body">
          <p class="ws-step-title">${s.title}</p>
          ${d ? `<p class="ws-step-detail">${d}</p>` : ""}
        </div>
      </li>`;
  }).join("");

  // ── 額外區塊 ──────────────────────────────────────────────
  let extraHtml = "";

  if ((phase === "ready" || phase === "chart1_ready") && opts.summary) {
    const s = opts.summary;
    const success = evaluateRecognitionSuccess();
    extraHtml = `
      ${phase === "ready" ? recognitionSuccessHtml(success) : `<div class="recognition-success mid"><strong>圖一辨識已完成</strong><span>整體辨識成功率會在圖二辨識完成後，依公司重疊度與結構合理性估算。</span></div>`}
      <div class="ws-summary">
        <div><span class="ws-stat-val">${s.master_count ?? "—"}</span><span class="ws-stat-lbl">主表公司</span></div>
        <div><span class="ws-stat-val">${s.review_count ?? "—"}</span><span class="ws-stat-lbl">待確認</span></div>
        <div><span class="ws-stat-val">${s.candidate_count ?? "—"}</span><span class="ws-stat-lbl">候選</span></div>
      </div>
      <button class="ws-goto-btn" id="wsGotoBtn">前往總覽 →</button>`;
  }

  if (phase === "chart2_error") {
    const s = opts.summary || {};
    const diagnostics = opts.diagnostics || {};
    const lastError = diagnostics.last_error || {};
    extraHtml = `
      <div class="ws-summary">
        <div><span class="ws-stat-val">${s.master_count ?? "—"}</span><span class="ws-stat-lbl">主表公司</span></div>
        <div><span class="ws-stat-val">—</span><span class="ws-stat-lbl">待確認</span></div>
        <div><span class="ws-stat-val">—</span><span class="ws-stat-lbl">候選</span></div>
      </div>
      <div class="ws-error-msg">圖二辨識失敗。圖一結構已保存，可先查看主表，或重新上傳圖二補充資訊。</div>
      ${lastError.code ? `<div class="ws-warn-box">診斷：${lastError.code}。${lastError.message || ""}</div>` : ""}
      <label class="ws-retry-label">
        <span>重新上傳圖二</span>
        <input type="file" id="wsChart2Retry" accept="image/*" />
      </label>
      <button class="ws-goto-btn ws-goto-outline" id="wsGotoBtn">查看圖一結果 →</button>`;
  }

  if (phase === "cancelled") {
    extraHtml = `<div class="ws-error-msg">任務已取消。系統會在當前分塊完成後停止後續辨識，已避免額外 API 消耗。</div>`;
  }

  if (phase === "error" && opts.error) {
    extraHtml = `<div class="ws-error-msg">${opts.error}</div>`;
  }

  el.innerHTML = `
    <p class="workspace-eyebrow">工作進度</p>
    <ul class="ws-steps">${stepsHtml}</ul>
    ${extraHtml}
    ${
      ["processing", "chart1_ready", "processing_chart2"].includes(phase) && state.taskId
        ? `<button class="ws-goto-btn ws-goto-outline" id="wsCancelTaskBtn">取消任務</button>`
        : ""
    }`;

  // ── 事件綁定 ──────────────────────────────────────────────
  document.getElementById("wsGotoBtn")?.addEventListener("click", () => setView("overview"));
  document.getElementById("wsCancelTaskBtn")?.addEventListener("click", async () => {
    try {
      await cancelCurrentTask();
      renderWorkspace("cancelled");
    } catch (err) {
      renderWorkspace("error", { error: `取消失敗：${err.message}` });
    }
  });

  const retryInput = document.getElementById("wsChart2Retry");
  if (retryInput) {
    retryInput.addEventListener("change", async (e) => {
      const file = e.target.files[0];
      if (!file || !state.taskId) return;
      renderWorkspace("processing_chart2", { msg: "重新上傳圖二中…" });
      try {
        const fd = new FormData();
        fd.append("chart2", file);
        await apiPost(`/api/tasks/${state.taskId}/analyze-chart2`, fd, true);
        renderWorkspace("processing_chart2", { msg: "圖二辨識中…" });
        const task = await pollTask(state.taskId, (t) => {
          if (t.status === "ready") hydrateTask(t);
          if (t.status === "processing_chart2") renderWorkspace("processing_chart2", { progress: t.chart2_progress, summary: t.summary });
        });
        hydrateTask(task);
        if (task.status === "chart2_error") {
          renderWorkspace("chart2_error", { taskId: task.id, error: task.error, summary: task.summary, diagnostics: task.analysis_diagnostics });
        } else if (task.status === "cancelled" || task.status === "cancel_requested") {
          renderWorkspace("cancelled", { summary: task.summary });
        } else {
          renderWorkspace("ready", { summary: task.summary });
        }
      } catch (err) {
        renderWorkspace("error", { error: err.message });
      }
    });
  }
}

async function confirmChart2Match(taskId, taskSnapshot = null) {
  renderWorkspace("processing_chart2", { msg: "配對中…" });
  try {
    const result = await apiPost(`/api/tasks/${taskId}/confirm-chart2`, {
      task_snapshot: taskSnapshot || loadTaskSnapshot(taskId),
    });
    applyTaskRefresh(result);
    renderWorkspace("ready", { summary: result.summary });
    setView("overview");
  } catch (err) {
    const expired = String(err.message || "").includes("404");
    renderWorkspace("error", {
      error: expired ? "任務已過期或伺服器已重啟，請重新開始分析。" : `配對失敗：${err.message}`,
    });
  }
}

async function createTaskFromUpload(onStatus) {
  if (!state.chart1File || !state.chart2File) {
    throw new Error("請先重新選擇圖一與圖二，再開始分析。");
  }
  const blockers = getUploadBlockers();
  if (blockers.length) {
    throw new Error(`上傳條件未通過：${blockers.join("；")}`);
  }

  const formData = new FormData();
  formData.append("task_name", elements.taskNameInput.value.trim());
  formData.append("chart1", state.chart1File);
  formData.append("chart2", state.chart2File);

  renderWorkspace("uploading");
  if (onStatus) onStatus("上傳中…");
  const initTask = await apiPost("/api/tasks/analyze", formData, true);

  renderWorkspace("processing", { msg: "圖一辨識中…" });
  if (onStatus) onStatus("AI 辨識中，請稍候…");

  try {
    const task = await pollTask(initTask.id, (t) => {
      if (t.status === "chart1_ready") {
        hydrateTask(t);
        renderWorkspace("chart1_ready", { summary: t.summary });
      } else if (t.status === "processing_chart2") {
        hydrateTask(t);
        renderWorkspace("processing_chart2", { progress: t.chart2_progress, summary: t.summary });
      } else if (t.status === "chart2_ocr_done") {
        hydrateTask(t);
        renderWorkspace("chart2_confirm", { task: t });
      } else {
        const secs = Math.round((Date.now() - _analysisStart) / 1000);
        renderWorkspace("processing", { msg: `辨識中… 已等候 ${secs} 秒` });
      }
      if (onStatus) onStatus(t.status);
    });

    if (task.status === "chart2_ocr_done") {
      // 圖二 OCR 完成：顯示確認畫面，不自動 merge
      renderWorkspace("chart2_confirm", { task });
      return;
    }

    hydrateTask(task);
    if (task.status === "chart2_error") {
      renderWorkspace("chart2_error", { taskId: task.id, error: task.error, summary: task.summary, diagnostics: task.analysis_diagnostics });
    } else if (task.status === "cancelled" || task.status === "cancel_requested") {
      renderWorkspace("cancelled", { summary: task.summary });
    } else {
      renderWorkspace("ready", { summary: task.summary });
    }
  } catch (err) {
    renderWorkspace("error", { error: err.message });
    throw err;
  }
}

async function createManualTask() {
  const taskName = elements.taskNameInput.value.trim();
  if (!taskName) {
    throw new Error("請先填寫集團名稱。");
  }
  const rootName = (elements.manualRootNameInput?.value || "").trim();
  const templateKey = (elements.manualTemplateSelect?.value || "blank").trim();
  renderWorkspace("uploading");
  const result = await apiPost("/api/tasks/create-manual", {
    task_name: taskName,
    root_company_name: rootName,
    template_key: templateKey,
  });
  const task = result.task || result;
  hydrateTask(task);
  renderWorkspace("ready", { summary: task.summary || {} });
  setView("results");
}

let _analysisStart = 0;

function exportWorkbook() {
  if (!window.XLSX) {
    alert("目前無法載入匯出元件，請稍後再試。");
    return;
  }

  const summaryRows = [
    ["集團名稱", state.taskName || "未命名集團"],
    ["任務 ID", state.taskId],
    ["主表公司數", state.masterRows.length],
    ["待確認數", state.reviewRows.length],
    ["圖二新增候選數", state.candidateRows.length],
  ];
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(summaryRows), "總覽");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(state.masterRows), "主表");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(state.reviewRows), "待確認");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(state.candidateRows), "新增候選");
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.json_to_sheet(
      Object.entries(state.reviewDecisions).map(([key, value]) => ({ key, ...value })),
    ),
    "人工確認紀錄",
  );
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.json_to_sheet(
      Object.entries(state.candidateDecisions).map(([key, value]) => ({ key, ...value })),
    ),
    "新增候選決策",
  );
  XLSX.writeFile(workbook, `${state.taskName || "股權圖整併審核結果"}.xlsx`);
}

function renderShareholderPanel() {
  if (!elements.shareholderList) return;

  if (!state.chartShareholders.length) {
    const showEmpty = !state.chartExternalEntities.length;
    elements.shareholderList.innerHTML = showEmpty
      ? `<span class="shareholder-empty">尚未加入上層股東</span>`
      : "";
    return;
  }

  const byId = {};
  state.masterRows.forEach((row) => { byId[row.node_id] = row; });
  elements.shareholderList.innerHTML = state.chartShareholders.map((holder) => {
    return `
      <article class="shareholder-chip">
        <span><strong>${holder.name}</strong></span>
        <button type="button" data-shareholder-edit="${holder.id}" title="編輯股東">✎</button>
        <button type="button" data-shareholder-delete="${holder.id}" title="移除此股東">×</button>
      </article>
    `;
  }).join("");
  elements.shareholderList.querySelectorAll("[data-shareholder-edit]").forEach((button) => {
    button.addEventListener("click", () => {
      const target = state.chartShareholders.find((holder) => holder.id === button.dataset.shareholderEdit);
      if (target) openShareholderModal(target);
    });
  });
  elements.shareholderList.querySelectorAll("[data-shareholder-delete]").forEach((button) => {
    button.addEventListener("click", () => {
      deleteChartShareholder(button.dataset.shareholderDelete).catch((error) => alert(`移除失敗：${error.message}`));
    });
  });
}

async function saveChartShareholders(nextShareholders) {
  const payload = await apiPost(`/api/tasks/${state.taskId}/chart-shareholders`, {
    chart_shareholders: nextShareholders,
  });
  applyTaskRefresh(payload);
  renderChart();
}

function renderExternalEntityPanel() {
  if (!elements.externalEntityList) return;
  if (!state.chartExternalEntities.length) {
    const showEmpty = !state.chartShareholders.length;
    elements.externalEntityList.innerHTML = showEmpty
      ? `<span class="external-empty">尚未加入集團外主體</span>`
      : "";
    return;
  }
  const byId = {};
  state.masterRows.forEach((row) => { byId[row.node_id] = row; });
  elements.externalEntityList.innerHTML = state.chartExternalEntities.map((entity) => {
    const names = normalizeExternalLayerNames(entity);
    const displayName = String(entity.name || "").trim() || names[0] || "集團外主體";
    const levelText = Number(entity.levels) >= 2 ? ` · ${Number(entity.levels)}層` : "";
    const fixedText = entity.placement_mode === "fixed" ? " · 已固定" : "";
    return `
      <article class="external-chip">
        <span><strong>${svgEscape(displayName)}</strong>${levelText}${fixedText}</span>
        <button type="button" data-external-auto="${entity.id}" title="恢復自動擺位">↺</button>
        <button type="button" data-external-edit="${entity.id}" title="編輯主體">✎</button>
        <button type="button" data-external-delete="${entity.id}" title="移除此主體">×</button>
      </article>`;
  }).join("");
  elements.externalEntityList.querySelectorAll("[data-external-auto]").forEach((button) => {
    button.addEventListener("click", () => {
      restoreExternalGroupAuto(button.dataset.externalAuto).catch((error) => alert(`恢復自動失敗：${error.message}`));
    });
  });
  elements.externalEntityList.querySelectorAll("[data-external-edit]").forEach((button) => {
    button.addEventListener("click", () => {
      const target = state.chartExternalEntities.find((entity) => entity.id === button.dataset.externalEdit);
      if (target) openExternalEntityModal(target);
    });
  });
  elements.externalEntityList.querySelectorAll("[data-external-delete]").forEach((button) => {
    button.addEventListener("click", () => {
      deleteExternalEntity(button.dataset.externalDelete).catch((error) => alert(`移除失敗：${error.message}`));
    });
  });
}

async function saveExternalEntities(nextEntities) {
  const payload = await apiPost(`/api/tasks/${state.taskId}/chart-external-entities`, {
    chart_external_entities: nextEntities,
  });
  state.externalLayoutNeedsFit = true;
  applyTaskRefresh(payload);
  renderChart();
}

async function addExternalEntityWithPayload(form) {
  const name = String(form?.name || "").trim();
  if (!name) return;
  const levels = Math.max(2, Math.min(4, Number(form.levels) || 2));
  const layerNames = normalizeExternalLayerNames({ name, levels, layer_names: form.layer_names || [] });
  const layerShares = normalizeExternalLayerShares({ levels, layer_shares: form.layer_shares || [] });
  const next = [
    ...state.chartExternalEntities,
    {
      id: makeShareholderId(),
      name,
      group: String(form.group || "").trim() || "集團外架構",
      target_node_id: form.target_node_id || "",
      share: String(form.share || "").trim(),
      levels,
      layer_names: layerNames,
      layer_shares: layerShares,
      placement_mode: "auto",
      manual_x: null,
      manual_y: null,
      note: "",
      external_scale: normalizeExternalScale(form.external_scale),
    },
  ];
  await saveExternalEntities(next);
}

async function updateExternalEntity(id, form) {
  if (!state.taskId || !id) return;
  const next = state.chartExternalEntities.map((entity) => {
    if (entity.id !== id) return entity;
    const levels = Math.max(2, Math.min(4, Number(form.levels) || entity.levels || 2));
    const layerNames = normalizeExternalLayerNames({
      ...entity,
      levels,
      layer_names: form.layer_names || entity.layer_names || [],
      name: String(form.name || "").trim() || entity.name,
    });
    const layerShares = normalizeExternalLayerShares({
      ...entity,
      levels,
      layer_shares: form.layer_shares || entity.layer_shares || [],
    });
    return {
      ...entity,
      name: String(form.name || "").trim() || entity.name,
      group: String(form.group || "").trim() || "集團外架構",
      target_node_id: form.target_node_id || "",
      share: String(form.share || "").trim(),
      levels,
      layer_names: layerNames,
      layer_shares: layerShares,
      placement_mode: entity.placement_mode === "fixed" ? "fixed" : "auto",
      manual_x: Number.isFinite(Number(entity.manual_x)) ? Number(entity.manual_x) : null,
      manual_y: Number.isFinite(Number(entity.manual_y)) ? Number(entity.manual_y) : null,
      external_scale: normalizeExternalScale(form.external_scale ?? entity.external_scale),
    };
  });
  await saveExternalEntities(next);
}

async function setExternalGroupPlacementByEntity(entityId, placement = {}) {
  if (!state.taskId || !entityId) return;
  const source = state.chartExternalEntities.find((entity) => entity.id === entityId);
  if (!source) return;
  const groupName = String(source.group || "").trim() || "集團外架構";
  const mode = placement.mode === "fixed" ? "fixed" : "auto";
  const next = state.chartExternalEntities.map((entity) => {
    const entityGroup = String(entity.group || "").trim() || "集團外架構";
    if (entityGroup !== groupName) return entity;
    return {
      ...entity,
      placement_mode: mode,
      manual_x: mode === "fixed" && Number.isFinite(Number(placement.manual_x)) ? Number(placement.manual_x) : null,
      manual_y: mode === "fixed" && Number.isFinite(Number(placement.manual_y)) ? Number(placement.manual_y) : null,
    };
  });
  await saveExternalEntities(next);
}

async function setExternalGroupPlacementByGroupName(groupName, placement = {}) {
  if (!state.taskId) return;
  const key = String(groupName || "").trim();
  if (!key) return;
  const mode = placement.mode === "fixed" ? "fixed" : "auto";
  const next = state.chartExternalEntities.map((entity) => {
    const entityGroup = String(entity.group || "").trim() || "集團外架構";
    if (entityGroup !== key) return entity;
    return {
      ...entity,
      placement_mode: mode,
      manual_x: mode === "fixed" && Number.isFinite(Number(placement.manual_x)) ? Number(placement.manual_x) : null,
      manual_y: mode === "fixed" && Number.isFinite(Number(placement.manual_y)) ? Number(placement.manual_y) : null,
    };
  });
  await saveExternalEntities(next);
}

async function restoreExternalGroupAuto(entityId) {
  await setExternalGroupPlacementByEntity(entityId, { mode: "auto" });
}

async function deleteExternalEntity(id) {
  if (!id) return;
  await saveExternalEntities(state.chartExternalEntities.filter((entity) => entity.id !== id));
}

async function addChartShareholderWithPayload(form) {
  if (!state.taskId) return;
  const name = String(form?.name || "").trim();
  const targetNodeId = form?.target_node_id || "";
  if (!name || !targetNodeId) {
    return;
  }
  const next = [
    ...state.chartShareholders,
    {
      id: makeShareholderId(),
      name,
      type: form.type || "company",
      share: String(form.share || "").trim(),
      target_node_id: targetNodeId,
      note: "",
    },
  ];
  await saveChartShareholders(next);
}

async function updateChartShareholder(id, form) {
  if (!state.taskId || !id) return;
  const next = state.chartShareholders.map((holder) => {
    if (holder.id !== id) return holder;
    return {
      ...holder,
      name: String(form.name || "").trim() || holder.name,
      type: form.type || holder.type || "company",
      share: String(form.share || "").trim(),
      target_node_id: form.target_node_id || holder.target_node_id,
    };
  });
  await saveChartShareholders(next);
}

async function deleteChartShareholder(id) {
  if (!state.taskId || !id) return;
  await saveChartShareholders(state.chartShareholders.filter((holder) => holder.id !== id));
}

function closeEntityModal() {
  document.getElementById("entityEditModal")?.remove();
}

function openShareholderModal(existing = null) {
  if (!state.taskId) return;
  const targets = getTopLevelCompanyRows();
  if (!targets.length) {
    alert("目前沒有可選的一級子公司。");
    return;
  }
  closeEntityModal();
  const modal = document.createElement("div");
  modal.id = "entityEditModal";
  modal.className = "entity-modal-backdrop";
  modal.innerHTML = `
    <div class="entity-modal" role="dialog" aria-modal="true" aria-labelledby="entityModalTitle">
      <h3 id="entityModalTitle">${existing ? "編輯上層股東" : "新增上層股東"}</h3>
      <label class="field"><span>名稱</span><input id="entityNameInput" type="text" placeholder="公司或個人名稱" value="${svgEscape(existing?.name || "")}" /></label>
      <label class="field"><span>類型</span>
        <select id="entityTypeSelect">
          <option value="company" ${existing?.type !== "person" ? "selected" : ""}>公司</option>
          <option value="person" ${existing?.type === "person" ? "selected" : ""}>個人</option>
        </select>
      </label>
      <label class="field"><span>持股</span><input id="entityShareInput" type="text" placeholder="例如：60%" value="${svgEscape(existing?.share || "")}" /></label>
      <label class="field"><span>投資對象</span>
        <select id="entityTargetSelect">${targets.map((row) => `<option value="${row.node_id}" ${existing?.target_node_id === row.node_id ? "selected" : ""}>${svgEscape(row.canonical_name || row.chart1_name || "未命名公司")}</option>`).join("")}</select>
      </label>
      <div class="detail-actions">
        <button class="ghost-btn" id="entityCancelBtn" type="button">取消</button>
        <button class="primary-btn" id="entitySubmitBtn" type="button">儲存</button>
      </div>
    </div>`;
  document.body.appendChild(modal);
  const nameInput = modal.querySelector("#entityNameInput");
  nameInput?.focus();
  modal.querySelector("#entityCancelBtn")?.addEventListener("click", closeEntityModal);
  modal.addEventListener("click", (event) => {
    if (event.target === modal) closeEntityModal();
  });
  modal.querySelector("#entitySubmitBtn")?.addEventListener("click", async () => {
    const name = modal.querySelector("#entityNameInput")?.value.trim() || "";
    const type = modal.querySelector("#entityTypeSelect")?.value || "company";
    const share = modal.querySelector("#entityShareInput")?.value.trim() || "";
    const target_node_id = modal.querySelector("#entityTargetSelect")?.value || "";
    if (!name || !target_node_id) return;
    if (existing?.id) {
      await updateChartShareholder(existing.id, { name, type, share, target_node_id });
    } else {
      await addChartShareholderWithPayload({ name, type, share, target_node_id });
    }
    closeEntityModal();
  });
}

function openExternalEntityModal(existing = null) {
  if (!state.taskId) return;
  const baseLevels = Math.max(2, Math.min(4, Number(existing?.levels) || 2));
  const baseNames = normalizeExternalLayerNames(existing || { levels: baseLevels, name: existing?.name || "" });
  const baseShares = normalizeExternalLayerShares(existing || { levels: baseLevels });
  closeEntityModal();
  const modal = document.createElement("div");
  modal.id = "entityEditModal";
  modal.className = "entity-modal-backdrop";
  modal.innerHTML = `
    <div class="entity-modal" role="dialog" aria-modal="true" aria-labelledby="entityModalTitle">
      <h3 id="entityModalTitle">${existing ? "編輯集團外主體" : "新增集團外主體"}</h3>
      <label class="field"><span>層級</span>
        <select id="entityLevelsSelect">
          <option value="2" ${baseLevels === 2 ? "selected" : ""}>2 層</option>
          <option value="3" ${baseLevels === 3 ? "selected" : ""}>3 層</option>
          <option value="4" ${baseLevels === 4 ? "selected" : ""}>4 層</option>
        </select>
      </label>
      <label class="field"><span>縮放（%）</span><input id="entityScaleInput" type="number" min="70" max="130" step="5" value="${normalizeExternalScale(existing?.external_scale)}" /></label>
      <div id="externalLayersEditor" class="field"></div>
      <div class="detail-actions">
        <button class="ghost-btn" id="entityCancelBtn" type="button">取消</button>
        <button class="primary-btn" id="entitySubmitBtn" type="button">儲存</button>
      </div>
    </div>`;
  document.body.appendChild(modal);
  const levelsSelect = modal.querySelector("#entityLevelsSelect");
  levelsSelect?.focus();
  const layersEditor = modal.querySelector("#externalLayersEditor");
  let liveNames = [...baseNames];
  let liveShares = [...baseShares];
  const renderLayersEditor = () => {
    if (!layersEditor) return;
    liveNames = Array.from(modal.querySelectorAll("[data-layer-name]")).map((el) => el.value.trim()).concat(liveNames).slice(0, 4);
    liveShares = Array.from(modal.querySelectorAll("[data-layer-share]")).map((el) => el.value.trim()).concat(liveShares).slice(0, 3);
    const levels = Math.max(2, Math.min(4, Number(modal.querySelector("#entityLevelsSelect")?.value || "2")));
    layersEditor.innerHTML = `
      <span>層級內容</span>
      ${Array.from({ length: levels }).map((_, idx) => `
        <label class="field">
          <span>${idx === 0 ? "最上層" : `第${idx + 1}層`}</span>
          <input data-layer-name="${idx}" type="text" value="${svgEscape(liveNames[idx] || "")}" placeholder="輸入名稱" />
        </label>
        ${idx < levels - 1 ? `
        <label class="field">
          <span>${idx === 0 ? "最上層→第二層持股比率" : `第${idx + 1}層→第${idx + 2}層持股比率`}</span>
          <input data-layer-share="${idx}" type="text" value="${svgEscape(liveShares[idx] || "")}" placeholder="例如：100%" />
        </label>` : ""}
      `).join("")}
    `;
  };
  renderLayersEditor();
  modal.querySelector("#entityLevelsSelect")?.addEventListener("change", renderLayersEditor);
  modal.querySelector("#entityCancelBtn")?.addEventListener("click", closeEntityModal);
  modal.addEventListener("click", (event) => {
    if (event.target === modal) closeEntityModal();
  });
  modal.querySelector("#entitySubmitBtn")?.addEventListener("click", async () => {
    const levels = Number(modal.querySelector("#entityLevelsSelect")?.value || "2");
    const external_scale = normalizeExternalScale(modal.querySelector("#entityScaleInput")?.value || 100);
    const layer_names = Array.from(modal.querySelectorAll("[data-layer-name]")).map((el) => el.value.trim());
    const layer_shares = Array.from(modal.querySelectorAll("[data-layer-share]")).map((el) => el.value.trim());
    if (!layer_names[0]) return;
    const name = layer_names[0];
    const group = String(existing?.group || "集團外架構").trim() || "集團外架構";
    const target_node_id = existing?.target_node_id || "";
    const share = existing?.share || "";
    if (existing?.id) {
      await updateExternalEntity(existing.id, { name, levels, group, target_node_id, share, layer_names, layer_shares, external_scale });
    } else {
      await addExternalEntityWithPayload({ name, levels, group, target_node_id, share, layer_names, layer_shares, external_scale });
    }
    closeEntityModal();
  });
}

// ══════════════════════════════════════════════════════════════
// 股權架構圖
// ══════════════════════════════════════════════════════════════

const LEVEL_COLORS = ["#1e3a5f", "#1d4ed8", "#0891b2", "#0d9488", "#059669", "#d97706"];
const LEVEL_NAMES  = ["頂層主體", "一級子公司", "二級子公司", "三級子公司", "四級子公司", "五級以上"];
let _chart = null;
let _elk = null;
let _elkRenderSeq = 0;

// 節點尺寸
const NODE_W = 220;
const NODE_H = 110;
const ELK_NODE_W = 260;
const ELK_NODE_H = 116;
const CHART_PROFILES = {
  a4: {
    label: "A4 橫式",
    nodeW: 220,
    nodeH: 84,
    maxNodeW: 285,
    nameLen: 11,
    nameLines: 2,
    detailLines: 1,
    nodeSpacing: 34,
    layerSpacing: 72,
    pad: 40,
    nameFont: 12,
    detailFont: 9.8,
    edgeFont: 11,
    minWidth: 1280,
    minHeight: 900,
  },
  a3: {
    label: "A3 橫式",
    nodeW: 250,
    nodeH: 96,
    maxNodeW: 330,
    nameLen: 12,
    nameLines: 3,
    detailLines: 3,
    nodeSpacing: 52,
    layerSpacing: 92,
    pad: 48,
    nameFont: 13,
    detailFont: 10.5,
    edgeFont: 12,
    minWidth: 1600,
    minHeight: 1130,
  },
  paged: {
    label: "依一級子公司分頁",
    nodeW: 240,
    nodeH: 96,
    maxNodeW: 315,
    nameLen: 12,
    nameLines: 2,
    detailLines: 4,
    nodeSpacing: 42,
    layerSpacing: 82,
    pad: 44,
    nameFont: 12.5,
    detailFont: 10,
    edgeFont: 11,
    minWidth: 1280,
    minHeight: 900,
  },
};

function svgEscape(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function wrapName(name, maxLen = 12) {
  if (!name || name.length <= maxLen) return name || "";
  const lines = [];
  for (let i = 0; i < name.length; i += maxLen) lines.push(name.slice(i, i + maxLen));
  return lines.join("\n");
}

function wrapTextLines(text, maxLen = 12, maxLines = 3) {
  const value = String(text || "").trim();
  if (!value) return [];
  const lines = [];
  for (let i = 0; i < value.length && lines.length < maxLines; i += maxLen) {
    lines.push(value.slice(i, i + maxLen));
  }
  if (value.length > maxLen * maxLines) {
    lines[lines.length - 1] = `${lines[lines.length - 1].slice(0, -1)}...`;
  }
  return lines;
}

function getElk() {
  if (_elk) return _elk;
  if (!window.ELK) return null;
  _elk = new window.ELK({
    defaultLayoutOptions: {
      "elk.algorithm": "layered",
      "elk.direction": "DOWN",
      "elk.edgeRouting": "ORTHOGONAL",
      "elk.spacing.nodeNode": "54",
      "elk.layered.spacing.nodeNodeBetweenLayers": "92",
      "elk.layered.nodePlacement.strategy": "NETWORK_SIMPLEX",
      "elk.layered.crossingMinimization.strategy": "LAYER_SWEEP",
    },
  });
  return _elk;
}

function getChartProfile() {
  return CHART_PROFILES[state.chartMode] || CHART_PROFILES.a4;
}

function applyChartIntent(intent, { fromUser = false } = {}) {
  const next = intent || "presentation";
  state.chartIntent = next;
  if (next === "print_single") {
    state.chartMode = "a4";
    state.chartDepth = "2";
    state.chartDirection = "down";
    state.printFitToPage = true;
  } else if (next === "print_paged") {
    state.chartMode = "paged";
    state.chartDepth = "all";
    state.chartDirection = "down";
    state.printFitToPage = false;
  } else {
    state.chartMode = "a3";
    state.chartDepth = "all";
  }
  if (fromUser && elements.chartAdvancedPanel) {
    elements.chartAdvancedPanel.open = false;
  }
}

function syncChartModeButtons() {
  const isGraph = state.chartView === "graph";
  elements.chartViewButtons.forEach((button) => {
    button.classList.toggle("active", button.dataset.chartView === state.chartView);
  });
  elements.chartIntentButtons.forEach((button) => {
    button.classList.toggle("active", button.dataset.chartIntent === state.chartIntent);
  });
  elements.chartDirectionButtons.forEach((button) => {
    button.classList.toggle("active", button.dataset.chartDirection === state.chartDirection);
  });
  elements.chartStyleButtons.forEach((button) => {
    button.classList.toggle("active", button.dataset.chartStyle === state.chartStyle);
  });
  elements.chartModeButtons.forEach((button) => {
    button.classList.toggle("active", button.dataset.chartMode === state.chartMode);
  });
  elements.chartDepthButtons.forEach((button) => {
    button.classList.toggle("active", button.dataset.chartDepth === state.chartDepth);
  });
  if (elements.chartShowRootToggle) elements.chartShowRootToggle.checked = state.showGroupRoot;
  if (elements.hybridModeSelect) elements.hybridModeSelect.value = state.hybridMode;
  if (elements.hybridThresholdSelect) elements.hybridThresholdSelect.value = String(state.hybridThreshold);
  if (elements.chartFontScaleSelect) elements.chartFontScaleSelect.value = String(state.chartFontScale);
  if (elements.chartDensitySelect) elements.chartDensitySelect.value = state.chartDensity;
  if (elements.chartBranchPicker) {
    elements.chartBranchPicker.style.display = isGraph && state.chartMode === "paged" ? "" : "none";
  }
  elements.chartZoomButtons.forEach((button) => {
    button.style.display = isGraph ? "" : "none";
  });
  if (elements.chartZoomLabel) {
    elements.chartZoomLabel.style.display = isGraph ? "" : "none";
    elements.chartZoomLabel.textContent = `${Math.round(state.chartScale * 100)}%`;
  }
}

function getChartTitle() {
  const base = getGroupName();
  return base.includes("股權架構圖") ? base : `${base}股權架構圖`;
}

function getPrintTitle() {
  return (state.printTitle || getChartTitle()).trim();
}

function buildElkGraph(rows, profile = getChartProfile(), graphId = "root") {
  const validRows = rows.filter((r) => r.node_id);
  const byId = {};
  validRows.forEach((row) => { byId[row.node_id] = row; });
  const childrenByParent = {};
  validRows.forEach((row) => {
    if (!row.chart1_parent || !byId[row.chart1_parent]) return;
    if (!childrenByParent[row.chart1_parent]) childrenByParent[row.chart1_parent] = [];
    childrenByParent[row.chart1_parent].push(row);
  });

  const collapsedChildIds = new Set();
  const columnNodes = [];
  Object.entries(childrenByParent).forEach(([parentId, childRows]) => {
    const parent = byId[parentId];
    if (!parent) return;
    const parentLevel = Number(parent.chart1_level) || 0;
    if (parentLevel < 1) return;
    const leafChildren = childRows.filter((child) => {
      const isLeaf = !(childrenByParent[child.node_id] || []).length;
      const excluded = child.is_chart_shareholder || child.is_external_entity || child.is_external_group;
      return isLeaf && !excluded;
    });
    const threshold = Math.max(2, Number(state.hybridThreshold) || HYBRID_COLUMN_THRESHOLD);
    const forceOn = state.hybridMode === "on";
    const forceOff = state.hybridMode === "off";
    if (forceOff) return;
    if (!forceOn && leafChildren.length < threshold) return;
    const sortedLeaves = leafChildren.sort((a, b) => (a.sort_index || 0) - (b.sort_index || 0));
    sortedLeaves.forEach((child) => collapsedChildIds.add(child.node_id));
    const columnId = `COL_${parentId}`;
    const itemRows = sortedLeaves.map((child) => {
      const name = child.canonical_name || child.chart1_name || "—";
      const ratio = ratioPercentText(child.actual_controller_share || "");
      const primary = ratio ? `${name}（${ratio}）` : name;
      let secondary = "";
      let tertiary = "";
      if (state.chartDensity !== "compact") {
        const line2 = [];
        if (child.legal_representative) line2.push(`法代：${child.legal_representative}`);
        if (child.established_date) line2.push(`成立：${child.established_date}`);
        secondary = line2.join(" ｜ ");
        if (state.chartDensity === "full") {
          const line3 = [];
          if (child.registered_capital) line3.push(`資本：${formatCapital(child.registered_capital)}`);
          if (child.role_label) line3.push(`定位：${child.role_label}`);
          tertiary = line3.join(" ｜ ");
        }
      }
      return { primary, secondary, tertiary };
    });
    columnNodes.push({
      node_id: columnId,
      canonical_name: "下層公司清單",
      chart1_name: "下層公司清單",
      chart1_level: parentLevel + 1,
      chart1_parent: parentId,
      role_label: `共 ${sortedLeaves.length} 家`,
      chart_note: "",
      node_status: "enriched",
      is_hybrid_column: true,
      hybrid_items: itemRows,
      hybrid_count: sortedLeaves.length,
    });
  });

  const visibleRows = [
    ...validRows.filter((row) => !collapsedChildIds.has(row.node_id)),
    ...columnNodes,
  ];
  const visibleIds = new Set(visibleRows.map((r) => r.node_id));
  const ids = visibleIds;
  const shareholderEdges = validRows
    .filter((r) => r.is_chart_shareholder && r.shareholder_target && ids.has(r.shareholder_target))
    .map((r) => ({
      id: `edge_${r.node_id}_${r.shareholder_target}`,
      sources: [r.node_id],
      targets: [r.shareholder_target],
      ratio: r.shareholder_share || "",
    }));
  const externalRelationEdges = validRows
    .filter((r) => r.is_external_entity && r.external_link_target && ids.has(r.external_link_target))
    .map((r) => ({
      id: `edge_${r.node_id}_${r.external_link_target}`,
      sources: [r.node_id],
      targets: [r.external_link_target],
      ratio: r.external_link_share || "",
    }));

  return {
    id: graphId,
    layoutOptions: {
      "elk.algorithm": "layered",
      "elk.direction": state.chartDirection === "right" ? "RIGHT" : "DOWN",
      "elk.edgeRouting": "ORTHOGONAL",
      "elk.spacing.nodeNode": String(profile.nodeSpacing),
      "elk.layered.spacing.nodeNodeBetweenLayers": String(profile.layerSpacing),
      "elk.layered.nodePlacement.strategy": "NETWORK_SIMPLEX",
      "elk.layered.crossingMinimization.strategy": "LAYER_SWEEP",
    },
    children: visibleRows.map((r) => {
      const name = r.canonical_name || r.chart1_name || "—";
      if (r.is_hybrid_column) {
        const itemCount = (r.hybrid_items || []).length;
        const secondaryCount = (r.hybrid_items || []).filter((item) => typeof item === "object" && item?.secondary).length;
        const tertiaryCount = (r.hybrid_items || []).filter((item) => typeof item === "object" && item?.tertiary).length;
        const cappedCount = Math.min(itemCount, 22);
        const perItemH = 22;
        const perSecondaryH = 17;
        const perTertiaryH = 17;
        const height = Math.max(
          130,
          72
            + cappedCount * perItemH
            + Math.min(secondaryCount, 22) * perSecondaryH
            + Math.min(tertiaryCount, 22) * perTertiaryH
        );
        const width = Math.min(420, Math.max(250, profile.nodeW + 70));
        return { id: r.node_id, width, height, row: r };
      }
      const baseWidth = r.is_chart_shareholder ? Math.max(190, profile.nodeW - 34) : profile.nodeW;
      const level = Number(r.chart1_level) || 0;
      const adaptiveWidth = Math.max(164, Math.min(profile.maxNodeW, 156 + Math.ceil(name.length / 7) * 16));
      const width = level >= 2 ? Math.min(baseWidth, adaptiveWidth) : Math.max(baseWidth, adaptiveWidth);
      const height = r.is_chart_shareholder ? Math.max(72, profile.nodeH - 18) : profile.nodeH;
      if (r.is_external_entity || r.is_external_group) {
        const scale = normalizeExternalScale(r.external_scale) / 100;
        return {
          id: r.node_id,
          width: Math.round(width * scale),
          height: Math.round(height * scale),
          row: r,
        };
      }
      return { id: r.node_id, width, height, row: r };
    }),
    edges: [
      ...visibleRows
      .filter((r) => r.chart1_parent && ids.has(r.chart1_parent))
      .map((r) => ({
        id: `edge_${r.chart1_parent}_${r.node_id}`,
        sources: [r.chart1_parent],
        targets: [r.node_id],
        ratio: r.actual_controller_share || "",
      })),
      ...shareholderEdges,
      ...externalRelationEdges,
    ],
  };
}

function rebalanceLayoutSymmetry(layout) {
  if (!layout?.children?.length || state.chartDirection === "right") return layout;
  const nodesById = {};
  (layout.children || []).forEach((node) => { nodesById[node.id] = node; });
  const childIdsByParent = {};
  const incomingSourcesByTarget = {};
  (layout.edges || []).forEach((edge) => {
    const source = edge.sources?.[0];
    const target = edge.targets?.[0];
    if (!source || !target || !nodesById[source] || !nodesById[target]) return;
    const sourceRow = nodesById[source]?.row || {};
    const targetRow = nodesById[target]?.row || {};
    const isExternalToMainLink = Boolean(
      sourceRow.is_external_entity &&
      !targetRow.is_external_entity &&
      !targetRow.is_external_group
    );
    if (isExternalToMainLink) return;
    if (!childIdsByParent[source]) childIdsByParent[source] = [];
    childIdsByParent[source].push(target);
    if (!incomingSourcesByTarget[target]) incomingSourcesByTarget[target] = [];
    incomingSourcesByTarget[target].push(source);
  });

  const shiftSubtree = (rootId, dx) => {
    if (!dx) return;
    const stack = [rootId];
    const visited = new Set();
    while (stack.length) {
      const id = stack.pop();
      if (visited.has(id)) continue;
      visited.add(id);
      const node = nodesById[id];
      if (node) node.x = (node.x || 0) + dx;
      (childIdsByParent[id] || []).forEach((childId) => stack.push(childId));
    }
  };

  Object.keys(childIdsByParent)
    .map((id) => nodesById[id])
    .filter(Boolean)
    .sort((a, b) => (a.y || 0) - (b.y || 0))
    .forEach((parent) => {
      const childNodes = (childIdsByParent[parent.id] || [])
        .map((id) => nodesById[id])
        .filter(Boolean);
      if (childNodes.length < 2) return;
      const minLeft = Math.min(...childNodes.map((n) => n.x || 0));
      const maxRight = Math.max(...childNodes.map((n) => (n.x || 0) + (n.width || 0)));
      const groupCenter = (minLeft + maxRight) / 2;
      const parentCenter = (parent.x || 0) + (parent.width || 0) / 2;
      const dx = Math.max(-120, Math.min(120, parentCenter - groupCenter));
      if (Math.abs(dx) < 6) return;
      childNodes.forEach((child) => shiftSubtree(child.id, dx));
    });

  // 上層股東：以目標公司中心做對稱排列
  Object.entries(incomingSourcesByTarget).forEach(([targetId, sourceIds]) => {
    const target = nodesById[targetId];
    if (!target) return;
    const shareholderNodes = sourceIds
      .map((id) => nodesById[id])
      .filter((node) => node?.row?.is_chart_shareholder);
    if (!shareholderNodes.length) return;
    shareholderNodes.sort((a, b) => (a.x || 0) - (b.x || 0));
    const gap = 18;
    const totalWidth = shareholderNodes.reduce((sum, node) => sum + (node.width || 0), 0) + gap * Math.max(0, shareholderNodes.length - 1);
    const targetCenter = (target.x || 0) + (target.width || 0) / 2;
    let cursorLeft = targetCenter - totalWidth / 2;
    shareholderNodes.forEach((node) => {
      const desiredCenter = cursorLeft + (node.width || 0) / 2;
      const currentCenter = (node.x || 0) + (node.width || 0) / 2;
      const dx = desiredCenter - currentCenter;
      if (Math.abs(dx) > 1) shiftSubtree(node.id, dx);
      cursorLeft += (node.width || 0) + gap;
    });
  });

  // 只有一個下一層子節點：強制置中對齊父節點
  Object.entries(childIdsByParent).forEach(([parentId, childIds]) => {
    if (childIds.length !== 1) return;
    const parent = nodesById[parentId];
    const child = nodesById[childIds[0]];
    if (!parent || !child) return;
    const parentCenter = (parent.x || 0) + (parent.width || 0) / 2;
    const childCenter = (child.x || 0) + (child.width || 0) / 2;
    const dx = parentCenter - childCenter;
    if (Math.abs(dx) > 1) shiftSubtree(child.id, dx);
  });

  // 第二段：防碰撞，確保同層節點不重疊
  const MIN_GAP = 22;
  const levelGroups = new Map();
  (layout.children || []).forEach((node) => {
    const key = String(node.row?.chart1_level ?? -1);
    if (!levelGroups.has(key)) levelGroups.set(key, []);
    levelGroups.get(key).push(node);
  });
  levelGroups.forEach((nodes) => {
    nodes.sort((a, b) => (a.x || 0) - (b.x || 0));
    for (let i = 1; i < nodes.length; i += 1) {
      const prev = nodes[i - 1];
      const curr = nodes[i];
      const prevRight = (prev.x || 0) + (prev.width || 0);
      const currLeft = curr.x || 0;
      const overlap = prevRight + MIN_GAP - currLeft;
      if (overlap > 0) {
        shiftSubtree(curr.id, overlap);
      }
    }
  });

  // 集團外主體：群組內鏈條固定為由上到下垂直主幹
  const externalVerticalGroups = new Map();
  (layout.children || []).forEach((node) => {
    const row = node?.row || {};
    const groupName = String(row.external_group || row.external_group_name || "").trim();
    if (!groupName) return;
    if (!externalVerticalGroups.has(groupName)) externalVerticalGroups.set(groupName, []);
    externalVerticalGroups.get(groupName).push(node);
  });
  externalVerticalGroups.forEach((nodes) => {
    if (nodes.length < 2) return;
    const chain = nodes
      .filter((node) => node?.row?.is_external_group || node?.row?.is_external_entity)
      .sort((a, b) => (Number(a.row?.chart1_level) || 0) - (Number(b.row?.chart1_level) || 0));
    if (chain.length < 2) return;
    const anchorX = (chain[0].x || 0) + (chain[0].width || 0) / 2;
    const verticalGap = 34;
    let cursorY = chain[0].y || 0;
    chain.forEach((node, idx) => {
      const desiredCenterX = anchorX;
      const currentCenterX = (node.x || 0) + (node.width || 0) / 2;
      const dx = desiredCenterX - currentCenterX;
      if (Math.abs(dx) > 0.5) shiftSubtree(node.id, dx);
      if (idx > 0) {
        const minY = cursorY + (chain[idx - 1].height || 0) + verticalGap;
        if ((node.y || 0) < minY) {
          const dy = minY - (node.y || 0);
          const stack = [node.id];
          const visited = new Set();
          while (stack.length) {
            const id = stack.pop();
            if (visited.has(id)) continue;
            visited.add(id);
            const targetNode = nodesById[id];
            if (targetNode) targetNode.y = (targetNode.y || 0) + dy;
            (childIdsByParent[id] || []).forEach((childId) => stack.push(childId));
          }
        }
      }
      cursorY = node.y || 0;
    });
  });

  // 集團外主體子圖：自動放在主幹圖空白角落（非固定左上）
  const externalRootNodes = (layout.children || []).filter((node) => node?.row?.is_external_group);
  if (externalRootNodes.length) {
    const collectSubtreeIds = (rootId, acc = new Set()) => {
      if (acc.has(rootId)) return acc;
      acc.add(rootId);
      (childIdsByParent[rootId] || []).forEach((childId) => collectSubtreeIds(childId, acc));
      return acc;
    };
    const bboxOfIds = (ids) => {
      const nodes = [...ids].map((id) => nodesById[id]).filter(Boolean);
      const minX = Math.min(...nodes.map((n) => n.x || 0));
      const minY = Math.min(...nodes.map((n) => n.y || 0));
      const maxX = Math.max(...nodes.map((n) => (n.x || 0) + (n.width || 0)));
      const maxY = Math.max(...nodes.map((n) => (n.y || 0) + (n.height || 0)));
      return { minX, minY, maxX, maxY, width: Math.max(1, maxX - minX), height: Math.max(1, maxY - minY) };
    };
    const overlapArea = (a, b) => {
      const w = Math.max(0, Math.min(a.maxX, b.maxX) - Math.max(a.minX, b.minX));
      const h = Math.max(0, Math.min(a.maxY, b.maxY) - Math.max(a.minY, b.minY));
      return w * h;
    };
    const nodeCenter = (node) => ({
      x: (node.x || 0) + (node.width || 0) / 2,
      y: (node.y || 0) + (node.height || 0) / 2,
    });

    const externalGroups = externalRootNodes.map((root) => {
      const ids = collectSubtreeIds(root.id);
      return { root, ids, bbox: bboxOfIds(ids) };
    });
    const externalIds = new Set(externalGroups.flatMap((g) => [...g.ids]));
    const mainNodes = (layout.children || []).filter((node) => !externalIds.has(node.id));
    if (mainNodes.length) {
      const mainBBox = bboxOfIds(new Set(mainNodes.map((n) => n.id)));
      const GAP = 70;
      const baseCandidates = [
        { x: mainBBox.minX - GAP, y: mainBBox.minY - GAP, name: "top-left" },
        { x: mainBBox.maxX + GAP, y: mainBBox.minY - GAP, name: "top-right" },
        { x: mainBBox.minX - GAP, y: mainBBox.maxY + GAP, name: "bottom-left" },
        { x: mainBBox.maxX + GAP, y: mainBBox.maxY + GAP, name: "bottom-right" },
      ];
      const shiftGroup = (group, dx, dy) => {
        if (!dx && !dy) return;
        const stack = [group.root.id];
        const visited = new Set();
        while (stack.length) {
          const id = stack.pop();
          if (visited.has(id)) continue;
          visited.add(id);
          const node = nodesById[id];
          if (node) {
            node.x = (node.x || 0) + dx;
            node.y = (node.y || 0) + dy;
          }
          (childIdsByParent[id] || []).forEach((childId) => stack.push(childId));
        }
      };

      const occupied = [mainBBox];
      externalGroups.forEach((group) => {
        const mode = group.root?.row?.external_placement_mode;
        const manualX = Number(group.root?.row?.external_manual_x);
        const manualY = Number(group.root?.row?.external_manual_y);
        if (mode !== "fixed" || !Number.isFinite(manualX) || !Number.isFinite(manualY)) return;
        const dx = manualX - (group.root.x || 0);
        const dy = manualY - (group.root.y || 0);
        shiftGroup(group, dx, dy);
        group.bbox = bboxOfIds(group.ids);
        occupied.push(group.bbox);
      });
      externalGroups.forEach((group, index) => {
        if (group.root?.row?.external_placement_mode === "fixed") return;
        const sourceEdge = (layout.edges || []).find((edge) => edge.sources?.[0] && group.ids.has(edge.sources[0]) && nodesById[edge.targets?.[0]]);
        const targetNode = sourceEdge ? nodesById[sourceEdge.targets[0]] : null;
        const targetCenter = targetNode ? nodeCenter(targetNode) : { x: mainBBox.maxX, y: mainBBox.maxY };
        const current = group.bbox;

        const candidates = [];
        for (let ring = 0; ring < 3; ring += 1) {
          const spread = ring * (Math.max(current.width, current.height) + 40);
          baseCandidates.forEach((base) => {
            const shiftX = /right/.test(base.name) ? spread : -spread;
            const shiftY = /bottom/.test(base.name) ? spread : -spread;
            const next = {
              minX: base.x + shiftX,
              minY: base.y + shiftY,
              maxX: base.x + shiftX + current.width,
              maxY: base.y + shiftY + current.height,
            };
            const overlap = occupied.reduce((sum, box) => sum + overlapArea(next, box), 0);
            const centerX = (next.minX + next.maxX) / 2;
            const centerY = (next.minY + next.maxY) / 2;
            const distance = Math.hypot(centerX - targetCenter.x, centerY - targetCenter.y);
            candidates.push({
              next,
              overlap,
              distance,
              score: overlap * 100000 + distance + ring * 180 + index * 15,
            });
          });
        }
        candidates.sort((a, b) => a.score - b.score);
        const best = candidates[0];
        if (!best) return;
        const dx = best.next.minX - current.minX;
        const dy = best.next.minY - current.minY;
        shiftGroup(group, dx, dy);
        group.bbox = bboxOfIds(group.ids);
        occupied.push(group.bbox);
      });
    }
  }

  const minX = Math.min(...(layout.children || []).map((n) => n.x || 0));
  if (minX < 0) {
    (layout.children || []).forEach((n) => { n.x = (n.x || 0) + Math.abs(minX) + 8; });
  }
  const minY = Math.min(...(layout.children || []).map((n) => n.y || 0));
  if (minY < 0) {
    (layout.children || []).forEach((n) => { n.y = (n.y || 0) + Math.abs(minY) + 8; });
  }
  layout.width = Math.max(...(layout.children || []).map((n) => (n.x || 0) + (n.width || 0))) + 16;
  layout.height = Math.max(...(layout.children || []).map((n) => (n.y || 0) + (n.height || 0))) + 16;
  return layout;
}

function buildPagedRowSets(rows) {
  const byId = {};
  rows.forEach((row) => {
    if (row.node_id) byId[row.node_id] = row;
  });

  const childrenByParent = {};
  rows.forEach((row) => {
    if (!row.chart1_parent || !byId[row.chart1_parent]) return;
    if (!childrenByParent[row.chart1_parent]) childrenByParent[row.chart1_parent] = [];
    childrenByParent[row.chart1_parent].push(row.node_id);
  });

  const roots = rows.filter((row) => !row.chart1_parent || !byId[row.chart1_parent]);
  const externalRootIds = new Set(roots.filter((row) => row.is_external_group).map((row) => row.node_id));
  const internalRoots = roots.filter((row) => !externalRootIds.has(row.node_id));
  if (!roots.length) return [{ title: "完整架構", rows }];

  const pages = [];
  function collectSubtree(id, seen = new Set()) {
    if (seen.has(id)) return [];
    seen.add(id);
    const current = byId[id] ? [byId[id]] : [];
    const childRows = (childrenByParent[id] || []).flatMap((childId) => collectSubtree(childId, seen));
    return [...current, ...childRows];
  }
  const externalRows = Array.from(externalRootIds).flatMap((rootId) => collectSubtree(rootId));

  if (!state.showGroupRoot) {
    internalRoots.forEach((root) => {
      const baseRows = collectSubtree(root.node_id);
      pages.push({
        id: root.node_id,
        title: root.canonical_name || root.chart1_name || "分頁",
        rows: externalRows.length ? [...baseRows, ...externalRows] : baseRows,
      });
    });
    return pages.length ? pages : [{ title: "完整架構", rows }];
  }

  internalRoots.forEach((root) => {
    const firstLevel = childrenByParent[root.node_id] || [];
    if (!firstLevel.length) {
      pages.push({
        id: root.node_id,
        title: root.canonical_name || root.chart1_name || "完整架構",
        rows: externalRows.length ? [root, ...externalRows] : [root],
      });
      return;
    }
    firstLevel.forEach((childId) => {
      const child = byId[childId];
      pages.push({
        id: childId,
        title: child?.canonical_name || child?.chart1_name || root.canonical_name || "分頁",
        rows: externalRows.length
          ? [root, ...collectSubtree(childId), ...externalRows]
          : [root, ...collectSubtree(childId)],
      });
    });
  });
  if (!internalRoots.length && externalRows.length) {
    pages.push({
      id: "external_entities_page",
      title: "集團外主體",
      rows: externalRows,
    });
  }

  return pages.length ? pages : [{ title: "完整架構", rows }];
}

function polylinePoints(edge) {
  const section = edge.sections?.[0];
  if (!section) return "";
  return [section.startPoint, ...(section.bendPoints || []), section.endPoint]
    .map((p) => `${p.x.toFixed(1)},${p.y.toFixed(1)}`)
    .join(" ");
}

function edgeLabelPosition(edge) {
  const section = edge.sections?.[0];
  if (!section) return null;
  const points = [section.startPoint, ...(section.bendPoints || []), section.endPoint];
  const mid = points[Math.floor(points.length / 2)];
  const next = points[Math.min(Math.floor(points.length / 2) + 1, points.length - 1)] || mid;
  return { x: (mid.x + next.x) / 2, y: (mid.y + next.y) / 2 };
}

function ratioPercentText(value) {
  const text = String(value || "").trim();
  if (!text) return "";
  const match = text.match(/-?\d+(?:\.\d+)?\s*%/);
  if (match) return match[0].replace(/\s+/g, "");
  const number = text.match(/-?\d+(?:\.\d+)?/);
  return number ? `${number[0]}%` : text;
}

function renderBranchEdges(layout, profile) {
  const drawPolyline = (points, { arrow = false } = {}) => `
    <polyline points="${points}" fill="none" stroke="#f8fafc" stroke-width="4.8" stroke-linecap="round" stroke-linejoin="round" />
    <polyline points="${points}" fill="none" stroke="#94a3b8" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" ${arrow ? 'marker-end="url(#elkArrow)"' : ""} />
  `;
  const nodesById = {};
  (layout.children || []).forEach((node) => {
    nodesById[node.id] = node;
  });

  const edgesByParent = {};
  (layout.edges || []).forEach((edge) => {
    const source = edge.sources?.[0];
    const target = edge.targets?.[0];
    if (!source || !target || !nodesById[source] || !nodesById[target]) return;
    if (!edgesByParent[source]) edgesByParent[source] = [];
    edgesByParent[source].push({ edge, child: nodesById[target] });
  });

  return Object.entries(edgesByParent).map(([parentId, childEdges]) => {
    const parent = nodesById[parentId];
    if (!parent || !childEdges.length) return "";

    if (state.chartDirection === "right") {
      return renderRightBranchGroup(parent, childEdges, profile);
    }

    const parentX = (parent.x || 0) + parent.width / 2;
    const parentBottom = (parent.y || 0) + parent.height;
    const children = childEdges
      .map(({ edge, child }) => ({
        edge,
        child,
        x: (child.x || 0) + child.width / 2,
        top: child.y || 0,
      }))
      .sort((a, b) => a.x - b.x);

    if (children.length === 1) {
      const item = children[0];
      const midY = parentBottom + Math.max(28, Math.min(64, (item.top - parentBottom) * 0.5));
      const labelX = (parentX + item.x) / 2;
      return `
        <g class="elk-edge elk-edge-branch">
          ${drawPolyline(`${parentX.toFixed(1)},${parentBottom.toFixed(1)} ${parentX.toFixed(1)},${midY.toFixed(1)} ${item.x.toFixed(1)},${midY.toFixed(1)} ${item.x.toFixed(1)},${item.top.toFixed(1)}`, { arrow: true })}
          ${item.edge.ratio ? renderEdgeRatioLabel(labelX, midY - 12, item.edge.ratio, profile) : ""}
        </g>`;
    }

    const minChildTop = Math.min(...children.map((item) => item.top));
    const minX = Math.min(...children.map((item) => item.x));
    const maxX = Math.max(...children.map((item) => item.x));
    const naturalBusY = parentBottom + Math.max(32, Math.min(76, (minChildTop - parentBottom) * 0.45));
    const busY = Math.min(minChildTop - 34, naturalBusY);
    const busStartX = Math.min(minX, parentX);
    const busEndX = Math.max(maxX, parentX);

    const laneGap = 18;
    const centerNudge = 14;
    const lanes = children.map((item) => ({ ...item, laneX: item.x }));
    lanes.sort((a, b) => a.laneX - b.laneX);
    for (let i = 1; i < lanes.length; i += 1) {
      if (lanes[i].laneX - lanes[i - 1].laneX < laneGap) {
        lanes[i].laneX = lanes[i - 1].laneX + laneGap;
      }
    }
    const centerHits = lanes.filter((lane) => Math.abs(lane.laneX - parentX) < 8);
    if (centerHits.length) {
      centerHits.forEach((lane, idx) => {
        lane.laneX += (idx % 2 === 0 ? -1 : 1) * centerNudge;
      });
    }
    const drops = lanes.map((item) => {
      const laneX = item.laneX;
      const elbowY = item.top - 18;
      const points = Math.abs(laneX - item.x) < 1
        ? `${laneX.toFixed(1)},${busY.toFixed(1)} ${laneX.toFixed(1)},${item.top.toFixed(1)}`
        : `${laneX.toFixed(1)},${busY.toFixed(1)} ${laneX.toFixed(1)},${elbowY.toFixed(1)} ${item.x.toFixed(1)},${elbowY.toFixed(1)} ${item.x.toFixed(1)},${item.top.toFixed(1)}`;
      return `
      ${drawPolyline(points, { arrow: true })}
      ${item.edge.ratio ? renderEdgeRatioLabel(item.x, item.top - 15, item.edge.ratio, profile) : ""}`;
    }).join("");

    return `
      <g class="elk-edge elk-edge-branch">
        ${drawPolyline(`${parentX.toFixed(1)},${parentBottom.toFixed(1)} ${parentX.toFixed(1)},${busY.toFixed(1)}`)}
        ${drawPolyline(`${busStartX.toFixed(1)},${busY.toFixed(1)} ${busEndX.toFixed(1)},${busY.toFixed(1)}`)}
        ${drops}
      </g>`;
  }).join("");
}

function renderRightBranchGroup(parent, childEdges, profile) {
  const drawPolyline = (points, { arrow = false } = {}) => `
    <polyline points="${points}" fill="none" stroke="#f8fafc" stroke-width="4.8" stroke-linecap="round" stroke-linejoin="round" />
    <polyline points="${points}" fill="none" stroke="#94a3b8" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" ${arrow ? 'marker-end="url(#elkArrow)"' : ""} />
  `;
  const parentRight = (parent.x || 0) + parent.width;
  const parentY = (parent.y || 0) + parent.height / 2;
  const children = childEdges
    .map(({ edge, child }) => ({
      edge,
      child,
      left: child.x || 0,
      y: (child.y || 0) + child.height / 2,
    }))
    .sort((a, b) => a.y - b.y);

  if (children.length === 1) {
    const item = children[0];
    const midX = parentRight + Math.max(28, Math.min(72, (item.left - parentRight) * 0.5));
    const labelY = (parentY + item.y) / 2;
    return `
      <g class="elk-edge elk-edge-branch">
        ${drawPolyline(`${parentRight.toFixed(1)},${parentY.toFixed(1)} ${midX.toFixed(1)},${parentY.toFixed(1)} ${midX.toFixed(1)},${item.y.toFixed(1)} ${item.left.toFixed(1)},${item.y.toFixed(1)}`, { arrow: true })}
        ${item.edge.ratio ? renderEdgeRatioLabel(midX + 18, labelY, item.edge.ratio, profile) : ""}
      </g>`;
  }

  const minChildLeft = Math.min(...children.map((item) => item.left));
  const minY = Math.min(...children.map((item) => item.y));
  const maxY = Math.max(...children.map((item) => item.y));
  const naturalBusX = parentRight + Math.max(34, Math.min(82, (minChildLeft - parentRight) * 0.45));
  const busX = Math.min(minChildLeft - 42, naturalBusX);
  const busStartY = Math.min(minY, parentY);
  const busEndY = Math.max(maxY, parentY);
  const branches = children.map((item) => `
    ${drawPolyline(`${busX.toFixed(1)},${item.y.toFixed(1)} ${item.left.toFixed(1)},${item.y.toFixed(1)}`, { arrow: true })}
    ${item.edge.ratio ? renderEdgeRatioLabel(item.left - 36, item.y - 14, item.edge.ratio, profile) : ""}
  `).join("");

  return `
    <g class="elk-edge elk-edge-branch">
      ${drawPolyline(`${parentRight.toFixed(1)},${parentY.toFixed(1)} ${busX.toFixed(1)},${parentY.toFixed(1)}`)}
      ${drawPolyline(`${busX.toFixed(1)},${busStartY.toFixed(1)} ${busX.toFixed(1)},${busEndY.toFixed(1)}`)}
      ${branches}
    </g>`;
}

function renderEdgeRatioLabel(x, y, ratio, profile) {
  const label = ratioPercentText(ratio);
  if (!label) return "";
  return `
    <g transform="translate(${x.toFixed(1)}, ${y.toFixed(1)})">
      <text class="edge-ratio-label" text-anchor="middle" dominant-baseline="middle" font-size="${profile.edgeFont}" font-weight="800">${svgEscape(label)}</text>
    </g>`;
}

function renderElkSvg(layout, profile = getChartProfile(), opts = {}) {
  const pad = profile.pad;
  const width = Math.max(profile.minWidth, Math.ceil((layout.width || 1000) + pad * 2));
  const height = Math.max(profile.minHeight, Math.ceil((layout.height || 700) + pad * 2));
  const nodes = layout.children || [];
  const edges = layout.edges || [];
  const title = opts.title || "";
  const svgId = opts.id || "elkChartSvg";
  const titleSvg = title ? `
    <text x="${pad}" y="${pad - 14}" fill="#1e293b" font-size="16" font-weight="800">${svgEscape(title)}</text>
  ` : "";
  const watermarkSvg = `
    <text x="${(width - 14).toFixed(1)}" y="${(height - 12).toFixed(1)}" fill="#94a3b8" font-size="11" font-weight="600" text-anchor="end">本圖由AI助手生成</text>
  `;

  const edgeSvg = `<g transform="translate(${pad}, ${pad})">${renderBranchEdges(layout, profile)}</g>`;
  const externalFrameSvg = (() => {
    const groups = new Map();
    nodes.forEach((node) => {
      const row = node.row || {};
      const groupName = String(row.external_group || row.external_group_name || "").trim();
      if (!groupName) return;
      if (!groups.has(groupName)) groups.set(groupName, []);
      groups.get(groupName).push(node);
    });
    const frames = [];
    groups.forEach((groupNodes, groupName) => {
      if (!groupNodes.length) return;
      const minX = Math.min(...groupNodes.map((n) => (n.x || 0))) - 22;
      const minY = Math.min(...groupNodes.map((n) => (n.y || 0))) - 34;
      const maxX = Math.max(...groupNodes.map((n) => (n.x || 0) + (n.width || 0))) + 22;
      const maxY = Math.max(...groupNodes.map((n) => (n.y || 0) + (n.height || 0))) + 22;
      const rootNode = groupNodes.find((n) => n.row?.is_external_group) || groupNodes[0];
      frames.push(`
        <g class="external-group-frame" data-external-group="${svgEscape(groupName)}" data-root-x="${Number(rootNode.x || 0).toFixed(2)}" data-root-y="${Number(rootNode.y || 0).toFixed(2)}">
          <rect x="${minX.toFixed(1)}" y="${minY.toFixed(1)}" width="${Math.max(120, maxX - minX).toFixed(1)}" height="${Math.max(80, maxY - minY).toFixed(1)}" rx="10" />
          <text x="${(minX + 10).toFixed(1)}" y="${(minY + 18).toFixed(1)}">${svgEscape(groupName)}</text>
        </g>
      `);
    });
    return `<g transform="translate(${pad}, ${pad})">${frames.join("")}</g>`;
  })();

  const nodeSvg = nodes.map((node) => {
    const r = node.row || {};
    const fontScale = Math.max(0.85, Math.min(1.2, Number(state.chartFontScale || 100) / 100));
    const nameFont = Math.max(10, profile.nameFont * fontScale);
    const detailFont = Math.max(8.5, profile.detailFont * fontScale);
    if (r.is_hybrid_column) {
      const itemLines = (r.hybrid_items || []).slice(0, 22);
      const extraCount = Math.max(0, (r.hybrid_items || []).length - itemLines.length);
      const header = svgEscape(`${r.role_label || ""}`);
      const lineStartY = 48;
      let dyCursor = 16;
      const renderItem = itemLines.map((item) => {
        const primaryText = typeof item === "string" ? item : String(item?.primary || "");
        const secondaryText = typeof item === "string" ? "" : String(item?.secondary || "");
        const tertiaryText = typeof item === "string" ? "" : String(item?.tertiary || "");
        const primary = svgEscape(primaryText);
        const secondary = svgEscape(secondaryText);
        const tertiary = svgEscape(tertiaryText);
        const part = `
          <tspan x="14" dy="${dyCursor}">• ${primary}</tspan>
          ${secondary ? `<tspan x="28" dy="16" class="hybrid-subline">· ${secondary}</tspan>` : ""}
          ${tertiary ? `<tspan x="28" dy="16" class="hybrid-subline">· ${tertiary}</tspan>` : ""}
        `;
        dyCursor = tertiary ? 24 : (secondary ? 22 : 20);
        return part;
      }).join("");
      return `
      <g class="elk-node elk-node-hybrid-column" transform="translate(${(node.x || 0) + pad}, ${(node.y || 0) + pad})">
        <rect width="${node.width}" height="${node.height}" rx="8" fill="#ffffff" stroke="#64748b" stroke-width="1.6" stroke-dasharray="5 4" />
        <text x="${node.width / 2}" y="28" text-anchor="middle" fill="#0f172a" font-size="${Math.max(12, nameFont - 0.5)}" font-weight="800">下層公司清單</text>
        <text x="${node.width / 2}" y="44" text-anchor="middle" fill="#475569" font-size="${Math.max(10, detailFont)}" font-weight="700">${header}</text>
        <text x="14" y="${lineStartY}" fill="#334155" font-size="${Math.max(10, detailFont)}" font-weight="600">
          ${renderItem}
          ${extraCount > 0 ? `<tspan x="14" dy="20">• 其餘 ${extraCount} 家...</tspan>` : ""}
        </text>
      </g>`;
    }
    const level = Number(r.chart1_level) || 0;
    const color = LEVEL_COLORS[Math.min(level, LEVEL_COLORS.length - 1)];
    const isMono = state.chartStyle === "mono";
    const isShareholder = Boolean(r.is_chart_shareholder);
    const uncertain = r.node_status !== "enriched";
    const fill = isShareholder ? "#f8fafc" : (isMono ? "#ffffff" : color);
    const stroke = isShareholder ? "#64748b" : (isMono ? (uncertain ? "#f59e0b" : "#334155") : (uncertain ? "#fbbf24" : "rgba(255,255,255,0.35)"));
    const nameColor = isShareholder || isMono ? "#0f172a" : "#ffffff";
    const detailColor = isShareholder || isMono ? "#334155" : "rgba(255,255,255,0.92)";
    const nameMaxLines = (Number(r.chart1_level) || 0) >= 2 ? 2 : profile.nameLines;
    const nameLines = wrapTextLines(r.canonical_name || r.chart1_name || "—", profile.nameLen, nameMaxLines);
    const repCap = [
      r.legal_representative ? `法代：${r.legal_representative}` : "",
      r.registered_capital ? `資本：${formatCapital(r.registered_capital)}` : "",
    ].filter(Boolean).join("  ");
    const baseDetails = [
      isShareholder ? `上層股東：${shareholderTypeText(r.shareholder_type)}` : "",
      repCap,
      r.established_date ? `成立：${r.established_date}` : "",
      r.role_label ? `定位：${r.role_label}` : "",
      r.chart_note ? `備註：${r.chart_note}` : "",
    ].filter(Boolean);
    const baseLimit = state.chartDensity === "compact" ? 1 : state.chartDensity === "full" ? Math.max(4, profile.detailLines + 1) : profile.detailLines;
    const detailLimit = fontScale > 1.1 ? Math.max(1, baseLimit - 1) : baseLimit;
    const details = baseDetails.slice(0, detailLimit);
    const nameStart = profile.nodeH <= 96 ? 31 - (nameLines.length - 1) * 8 : 36 - (nameLines.length - 1) * 9;
    const detailStart = profile.nodeH - (details.length > 1 ? 35 : 26);
    const nodeClasses = ["elk-node"];
    if (r.is_external_group) nodeClasses.push("elk-node-external-group");
    const groupAttr = r.is_external_group ? `data-external-group="${svgEscape(String(r.external_group_name || r.canonical_name || ""))}"` : "";
    return `
      <g class="${nodeClasses.join(" ")}" data-node-id="${svgEscape(String(node.id || ""))}" data-layout-x="${Number(node.x || 0).toFixed(2)}" data-layout-y="${Number(node.y || 0).toFixed(2)}" ${groupAttr} transform="translate(${(node.x || 0) + pad}, ${(node.y || 0) + pad})">
        <rect width="${node.width}" height="${node.height}" rx="${isShareholder ? 18 : 4}" fill="${fill}" stroke="${stroke}" stroke-width="${isShareholder ? 1.8 : (uncertain ? 2.2 : 1.4)}" ${uncertain || isShareholder ? 'stroke-dasharray="7 4"' : ""} />
        <text x="${node.width / 2}" y="${nameStart}" text-anchor="middle" fill="${nameColor}" font-size="${nameFont}" font-weight="800">
          ${nameLines.map((line, i) => `<tspan x="${node.width / 2}" dy="${i === 0 ? 0 : 18}">${svgEscape(line)}</tspan>`).join("")}
        </text>
        <text x="${node.width / 2}" y="${detailStart}" text-anchor="middle" fill="${detailColor}" font-size="${detailFont}" font-weight="600">
          ${details.map((line, i) => `<tspan x="${node.width / 2}" dy="${i === 0 ? 0 : 15}">${svgEscape(line)}</tspan>`).join("")}
        </text>
        ${uncertain ? `<text x="${node.width - 14}" y="20" text-anchor="middle" fill="#f59e0b" font-size="15" font-weight="900">!</text>` : ""}
      </g>`;
  }).join("");

  return `
    <svg class="elk-svg" id="${svgId}" data-chart-mode="${state.chartMode}" xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" role="img" aria-label="股權架構圖">
      <defs>
        <marker id="elkArrow" viewBox="0 0 10 10" refX="8.5" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse">
          <path d="M 0 0 L 10 5 L 0 10 z" fill="#94a3b8"></path>
        </marker>
      </defs>
      <rect width="100%" height="100%" fill="#f8fafc"></rect>
      ${titleSvg}
      ${externalFrameSvg}
      ${edgeSvg}
      ${nodeSvg}
      ${watermarkSvg}
    </svg>`;
}

function renderChartViewport(markup) {
  return `
    <div class="chart-panzoom-shell" id="chartPanZoomShell">
      <div class="chart-panzoom-content" id="chartPanZoomContent">
        ${markup}
      </div>
    </div>`;
}

function applyChartViewport() {
  const content = document.getElementById("chartPanZoomContent");
  if (!content) return;
  content.style.transform = `translate(${state.chartPanX}px, ${state.chartPanY}px) scale(${state.chartScale})`;
  if (elements.chartZoomLabel) elements.chartZoomLabel.textContent = `${Math.round(state.chartScale * 100)}%`;
}

function setChartZoom(nextScale, anchor = null) {
  const previous = state.chartScale;
  const scale = clampNumber(nextScale, CHART_ZOOM_MIN, CHART_ZOOM_MAX);
  if (anchor) {
    state.chartPanX = anchor.x - ((anchor.x - state.chartPanX) * scale) / previous;
    state.chartPanY = anchor.y - ((anchor.y - state.chartPanY) * scale) / previous;
  }
  state.chartScale = scale;
  applyChartViewport();
}

function resetChartViewport(scale = 1) {
  state.chartScale = clampNumber(scale, CHART_ZOOM_MIN, CHART_ZOOM_MAX);
  state.chartPanX = 0;
  state.chartPanY = 0;
  applyChartViewport();
}

function fitChartToViewport() {
  const shell = document.getElementById("chartPanZoomShell");
  const svg = document.querySelector(".elk-svg");
  if (!shell || !svg) return;
  const width = Number(svg.getAttribute("width")) || svg.viewBox?.baseVal?.width || svg.scrollWidth || 1;
  const available = Math.max(shell.clientWidth - 48, 320);
  resetChartViewport(Math.min(1.25, available / width));
}

function bindChartViewport() {
  const shell = document.getElementById("chartPanZoomShell");
  if (!shell) return;
  let dragging = false;
  let startX = 0;
  let startY = 0;
  let originX = 0;
  let originY = 0;

  shell.addEventListener("pointerdown", (event) => {
    if (event.button !== 0) return;
    dragging = true;
    startX = event.clientX;
    startY = event.clientY;
    originX = state.chartPanX;
    originY = state.chartPanY;
    shell.setPointerCapture(event.pointerId);
    shell.classList.add("is-panning");
  });
  shell.addEventListener("pointermove", (event) => {
    if (!dragging) return;
    state.chartPanX = originX + event.clientX - startX;
    state.chartPanY = originY + event.clientY - startY;
    applyChartViewport();
  });
  shell.addEventListener("pointerup", (event) => {
    dragging = false;
    shell.releasePointerCapture(event.pointerId);
    shell.classList.remove("is-panning");
  });
  shell.addEventListener("wheel", (event) => {
    if (!event.ctrlKey && !event.metaKey) return;
    event.preventDefault();
    const rect = shell.getBoundingClientRect();
    const anchor = { x: event.clientX - rect.left, y: event.clientY - rect.top };
    setChartZoom(state.chartScale + (event.deltaY > 0 ? -0.08 : 0.08), anchor);
  }, { passive: false });
  applyChartViewport();
}

function bindExternalGroupDrag() {
  const shell = document.getElementById("chartPanZoomShell");
  if (!shell) return;
  const SNAP_DISTANCE = 20;
  let guideX = shell.querySelector(".external-snap-guide-x");
  let guideY = shell.querySelector(".external-snap-guide-y");
  if (!guideX) {
    guideX = document.createElement("div");
    guideX.className = "external-snap-guide external-snap-guide-x";
    shell.appendChild(guideX);
  }
  if (!guideY) {
    guideY = document.createElement("div");
    guideY.className = "external-snap-guide external-snap-guide-y";
    shell.appendChild(guideY);
  }
  const hideGuides = () => {
    guideX.classList.remove("is-visible");
    guideY.classList.remove("is-visible");
  };

  shell.querySelectorAll(".external-group-frame").forEach((groupNode) => {
    let dragging = null;
    let rafId = 0;
    const snapPoint = (x, y) => {
      if (!dragging) return { x, y };
      const profile = getChartProfile();
      const svg = shell.querySelector(".elk-svg");
      const centerX = svg ? ((Number(svg.getAttribute("width")) || 0) / 2 - profile.pad) : null;
      const peers = dragging.peerRoots || [];
      let nextX = x;
      let nextY = y;
      let snappedX = null;
      let snappedY = null;
      [centerX, ...peers.map((p) => p.x)].forEach((axisX) => {
        if (!Number.isFinite(axisX)) return;
        if (Math.abs(nextX - axisX) <= SNAP_DISTANCE) {
          nextX = axisX;
          snappedX = axisX;
        }
      });
      peers.forEach((p) => {
        if (!Number.isFinite(p.y)) return;
        if (Math.abs(nextY - p.y) <= SNAP_DISTANCE) {
          nextY = p.y;
          snappedY = p.y;
        }
      });
      if (snappedX !== null) {
        guideX.style.left = `${((snappedX + profile.pad) * state.chartScale + state.chartPanX).toFixed(1)}px`;
        guideX.classList.add("is-visible");
      } else {
        guideX.classList.remove("is-visible");
      }
      if (snappedY !== null) {
        guideY.style.top = `${((snappedY + profile.pad) * state.chartScale + state.chartPanY).toFixed(1)}px`;
        guideY.classList.add("is-visible");
      } else {
        guideY.classList.remove("is-visible");
      }
      return { x: nextX, y: nextY };
    };
    const flushDragTransform = () => {
      rafId = 0;
      if (!dragging) return;
      const dx = dragging.nextX - dragging.startLayoutX;
      const dy = dragging.nextY - dragging.startLayoutY;
      groupNode.style.transform = `translate3d(${dx.toFixed(1)}px, ${dy.toFixed(1)}px, 0)`;
    };
    groupNode.addEventListener("pointerdown", (event) => {
      if (event.button !== 0) return;
      event.stopPropagation();
      event.preventDefault();
      const groupName = String(groupNode.getAttribute("data-external-group") || "").trim();
      if (!groupName) return;
      const startLayoutX = Number(groupNode.getAttribute("data-root-x") || 0);
      const startLayoutY = Number(groupNode.getAttribute("data-root-y") || 0);
      dragging = {
        groupName,
        startClientX: event.clientX,
        startClientY: event.clientY,
        startLayoutX,
        startLayoutY,
        nextX: startLayoutX,
        nextY: startLayoutY,
        velocityX: 0,
        velocityY: 0,
        lastClientX: event.clientX,
        lastClientY: event.clientY,
        lastMoveAt: performance.now(),
        peerRoots: [...shell.querySelectorAll(".external-group-frame")]
          .filter((item) => item !== groupNode)
          .map((item) => ({
            x: Number(item.getAttribute("data-root-x") || 0),
            y: Number(item.getAttribute("data-root-y") || 0),
          })),
      };
      groupNode.setPointerCapture(event.pointerId);
      groupNode.classList.add("is-dragging");
      const hint = document.createElement("div");
      hint.id = "externalDragHint";
      hint.className = "external-drag-hint";
      hint.textContent = `拖曳中：${groupName}（放開後固定位置）`;
      shell.appendChild(hint);
      shell.classList.add("is-external-dragging");
    });

    groupNode.addEventListener("pointermove", (event) => {
      if (!dragging) return;
      const scale = Math.max(state.chartScale, 0.01);
      const dx = (event.clientX - dragging.startClientX) / scale;
      const dy = (event.clientY - dragging.startClientY) / scale;
      const targetX = dragging.startLayoutX + dx;
      const targetY = dragging.startLayoutY + dy;
      const snapped = snapPoint(targetX, targetY);
      dragging.nextX += (snapped.x - dragging.nextX) * 0.45;
      dragging.nextY += (snapped.y - dragging.nextY) * 0.45;
      const now = performance.now();
      const dt = Math.max(8, now - dragging.lastMoveAt);
      dragging.velocityX = ((event.clientX - dragging.lastClientX) / scale) / dt * 16;
      dragging.velocityY = ((event.clientY - dragging.lastClientY) / scale) / dt * 16;
      dragging.lastClientX = event.clientX;
      dragging.lastClientY = event.clientY;
      dragging.lastMoveAt = now;
      if (!rafId) rafId = requestAnimationFrame(flushDragTransform);
    });

    groupNode.addEventListener("pointerup", async () => {
      if (!dragging) return;
      const target = dragging;
      dragging = null;
      hideGuides();
      shell.classList.remove("is-external-dragging");
      document.getElementById("externalDragHint")?.remove();
      groupNode.classList.remove("is-dragging");
      if (rafId) cancelAnimationFrame(rafId);
      rafId = 0;
      if (!Number.isFinite(target.nextX) || !Number.isFinite(target.nextY)) {
        groupNode.style.transform = "";
        return;
      }
      let finalX = target.nextX;
      let finalY = target.nextY;
      let vx = target.velocityX;
      let vy = target.velocityY;
      await new Promise((resolve) => {
        const startedAt = performance.now();
        const step = () => {
          const elapsed = performance.now() - startedAt;
          if (elapsed > 240) return resolve();
          vx *= 0.9;
          vy *= 0.9;
          finalX += vx;
          finalY += vy;
          const snapped = snapPoint(finalX, finalY);
          finalX = snapped.x;
          finalY = snapped.y;
          const dx = finalX - target.startLayoutX;
          const dy = finalY - target.startLayoutY;
          groupNode.style.transform = `translate3d(${dx.toFixed(1)}px, ${dy.toFixed(1)}px, 0)`;
          if (Math.abs(vx) < 0.08 && Math.abs(vy) < 0.08) return resolve();
          requestAnimationFrame(step);
        };
        requestAnimationFrame(step);
      });
      groupNode.style.transform = "";
      hideGuides();
      await setExternalGroupPlacementByGroupName(target.groupName, {
        mode: "fixed",
        manual_x: finalX,
        manual_y: finalY,
      }).catch((error) => alert(`固定位置失敗：${error.message}`));
    });

    groupNode.addEventListener("pointercancel", () => {
      dragging = null;
      hideGuides();
      shell.classList.remove("is-external-dragging");
      document.getElementById("externalDragHint")?.remove();
      groupNode.classList.remove("is-dragging");
      groupNode.style.transform = "";
      if (rafId) cancelAnimationFrame(rafId);
      rafId = 0;
    });
  });
}

function syncBranchSelect(pages) {
  if (!elements.chartBranchSelect) return;
  const options = pages.map((page) => ({
    value: page.id || page.title,
    label: page.title,
  }));
  const validValues = new Set(options.map((option) => option.value));
  if (state.selectedBranchId !== "__all__" && !validValues.has(state.selectedBranchId)) {
    state.selectedBranchId = "__all__";
  }
  elements.chartBranchSelect.innerHTML = [
    `<option value="__all__">全部一級分支</option>`,
    ...options.map((option) => `<option value="${svgEscape(option.value)}">${svgEscape(option.label)}</option>`),
  ].join("");
  elements.chartBranchSelect.value = state.selectedBranchId;
}

async function renderElkChart() {
  const elk = getElk();
  const profile = getChartProfile();
  const chartRows = getChartRows();
  if (!elk) {
    elements.chartContainer.classList.add("chart-container-list");
    elements.chartContainer.classList.remove("chart-container-elk");
    renderListTree();
    elements.chartLayoutBadge.textContent = `${chartRows.length} 家公司 · ELK 未載入，使用條列樹狀`;
    return;
  }

  const seq = ++_elkRenderSeq;
  if (_chart) { _chart.dispose(); _chart = null; }
  elements.chartContainer.innerHTML = `<div class="elk-loading">${profile.label} 排版中...</div>`;

  try {
    if (state.chartMode === "paged") {
      const pages = buildPagedRowSets(chartRows);
      syncBranchSelect(pages);
      const visiblePages = state.selectedBranchId === "__all__"
        ? pages
        : pages.filter((page) => (page.id || page.title) === state.selectedBranchId);
      const svgs = [];
      for (let i = 0; i < visiblePages.length; i += 1) {
        const page = visiblePages[i];
        const sourceIndex = pages.findIndex((item) => (item.id || item.title) === (page.id || page.title));
        const graph = buildElkGraph(chartRowsWithShareholders(page.rows), profile, `page_${i + 1}`);
        if (!graph.children.length) continue;
        const layout = rebalanceLayoutSymmetry(await elk.layout(graph));
        const pageTitle = state.selectedBranchId === "__all__"
          ? `第 ${sourceIndex + 1} 頁 / ${pages.length}：${page.title}`
          : `分支：${page.title}`;
        svgs.push(`
          <article class="elk-page">
            <div class="elk-page-title">${svgEscape(pageTitle)}</div>
            ${renderElkSvg(layout, profile, { id: `elkChartSvg_${i + 1}` })}
          </article>
        `);
      }
      if (seq !== _elkRenderSeq) return;
      elements.chartContainer.innerHTML = svgs.length ? renderChartViewport(`<div id="elkPagedChart" class="elk-pages">${svgs.join("")}</div>`) : `<div class="elk-empty">沒有可顯示的公司資料</div>`;
      bindChartViewport();
      bindExternalGroupDrag();
      if (state.externalLayoutNeedsFit) {
        fitChartToViewport();
        state.externalLayoutNeedsFit = false;
      }
      return;
    }

    syncBranchSelect([]);
    const graph = buildElkGraph(chartRowsWithShareholders(chartRows), profile);
    if (!graph.children.length) {
      elements.chartContainer.innerHTML = `<div class="elk-empty">沒有可顯示的公司資料</div>`;
      return;
    }
    const layout = rebalanceLayoutSymmetry(await elk.layout(graph));
    if (seq !== _elkRenderSeq) return;
    elements.chartContainer.innerHTML = renderChartViewport(renderElkSvg(layout, profile));
    bindChartViewport();
    bindExternalGroupDrag();
    if (state.externalLayoutNeedsFit) {
      fitChartToViewport();
      state.externalLayoutNeedsFit = false;
    }
  } catch (error) {
    console.error(error);
    elements.chartContainer.classList.add("chart-container-list");
    elements.chartContainer.classList.remove("chart-container-elk");
    renderListTree();
    elements.chartLayoutBadge.textContent = `${chartRows.length} 家公司 · ELK 排版失敗，使用條列樹狀`;
  }
}

function buildEChartsTree(rows) {
  const byId = {};
  rows.forEach((r) => { byId[r.node_id] = { ...r, _children: [] }; });

  const roots = [];
  rows.forEach((r) => {
    const parent = r.chart1_parent && byId[r.chart1_parent];
    if (parent) parent._children.push(r.node_id);
    else roots.push(r.node_id);
  });

  function toNode(id) {
    const r = byId[id];
    if (!r) return null;
    const uncertain = r.node_status !== "enriched";
    const level = Number(r.chart1_level) || 0;
    const color = LEVEL_COLORS[Math.min(level, LEVEL_COLORS.length - 1)];

    // 公司名稱（最多兩行）
    const nameLine = wrapName(r.canonical_name || r.chart1_name || "—", 12);
    // 法代＋資本額同一行，成立日期另一行
    const repCap = [
      r.legal_representative ? `法代：${r.legal_representative}` : "",
      r.registered_capital   ? `資本：${formatCapital(r.registered_capital)}` : "",
    ].filter(Boolean).join("  ");

    const labelParts = [`{name|${uncertain ? "⚠ " : ""}${nameLine}}`];
    if (repCap)            labelParts.push(`{info|${repCap}}`);
    if (r.established_date) labelParts.push(`{info|成立：${r.established_date}}`);

    return {
      name: id,
      _row: r,
      label: { formatter: labelParts.join("\n") },
      itemStyle: {
        color,
        borderColor:  uncertain ? "#fbbf24" : "rgba(255,255,255,0.25)",
        borderWidth:  uncertain ? 3 : 1,
        borderType:   uncertain ? "dashed" : "solid",
        shadowColor:  "rgba(0,0,0,0.22)",
        shadowBlur:   10,
        shadowOffsetY: 3,
      },
      // 持股比例顯示在連線上
      edgeLabel: r.actual_controller_share ? {
        show: true,
        formatter: ratioPercentText(r.actual_controller_share),
        fontSize: 12,
        fontWeight: "bold",
        color: "#1e293b",
        textBorderColor: "#ffffff",
        textBorderWidth: 3,
      } : undefined,
      children: r._children.map(toNode).filter(Boolean),
    };
  }

  if (roots.length === 0) return null;
  if (roots.length === 1) return toNode(roots[0]);

  // 多個根：加一個隱形虛根
  return {
    name: "__root__",
    label: { show: false },
    itemStyle: { opacity: 0 },
    symbolSize: [0, 0],
    children: roots.map(toNode).filter(Boolean),
  };
}

// ── 條列式樹狀（>20 家） ──────────────────────────────────────
function renderListTree() {
  const container = elements.chartContainer;
  const chartRows = getChartRows();
  const roots = buildTree(chartRows);

  // 更新圖例
  const levels = [...new Set(chartRows.map((r) => Number(r.chart1_level) || 0))].sort();
  elements.chartLegend.innerHTML = [
    ...levels.map((l) => {
      const color = LEVEL_COLORS[Math.min(l, LEVEL_COLORS.length - 1)];
      return `<span class="legend-item"><span class="legend-dot" style="background:${color}"></span>${LEVEL_NAMES[l] || `第${l}層`}</span>`;
    }),
    `<span class="legend-item"><span class="legend-dot legend-dot-uncertain"></span>待確認</span>`,
  ].join("");

  let rows = [];

  function escapeAttr(value) {
    return String(value || "")
      .replaceAll("&", "&amp;")
      .replaceAll("\"", "&quot;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;");
  }

  function companyLinkByName(name = "") {
    const text = String(name || "").trim();
    if (!text) return "";
    if (/[協协]創數據技術股份有限公司/.test(text)) return "https://www.sharetronic.com/";
    if (/追覓|追觅|DREAME|Dreame/i.test(text)) return "https://www.dreame.tech/";
    if (/華碩|华硕|ASUS/i.test(text)) return "https://www.asus.com/";
    return `https://www.qcc.com/web/search?key=${encodeURIComponent(text)}`;
  }

  function walk(node, prefixLines, isLast) {
    const level = Number(node.chart1_level) || 0;
    const color = LEVEL_COLORS[Math.min(level, LEVEL_COLORS.length - 1)];
    const uncertain = node.node_status !== "enriched";

    // 前綴字元
    const isRoot = level === 0 && prefixLines.length === 0;
    const connector = isRoot ? "" : (isLast ? "└── " : "├── ");
    const prefix = prefixLines.join("") + connector;

    // 資訊欄
    const name = node.canonical_name || node.chart1_name || "—";
    const link = (level <= 1) ? companyLinkByName(name) : "";
    const attrs = [
      node.legal_representative ? `法代：${node.legal_representative}` : "",
      node.registered_capital   ? `資本：${formatCapital(node.registered_capital)}` : "",
      node.established_date     ? `成立：${node.established_date}` : "",
      node.actual_controller_share ? `持股：${node.actual_controller_share}` : "",
      node.role_label           ? `定位：${node.role_label}` : "",
      node.chart_note            ? `備註：${node.chart_note}` : "",
    ].filter(Boolean).join("｜");

    rows.push({ prefix, name, attrs, color, uncertain, isRoot, level, link });

    if (node.children?.length) {
      const childBase = isRoot ? "" : (isLast ? "    " : "│   ");
      node.children.forEach((child, i) => {
        const childIsLast = i === node.children.length - 1;
        walk(child, [...prefixLines, childBase], childIsLast);
      });
    }
  }

  roots.forEach((root, i) => walk(root, [], i === roots.length - 1));

  container.innerHTML = `
    <div class="list-tree" id="listTreeInner">
      ${rows.map((r) => `
        <div class="lt-row ${r.uncertain ? "lt-uncertain" : ""} ${r.isRoot ? "lt-root" : ""}">
          <span class="lt-prefix">${r.isRoot ? "" : r.prefix}</span>
          ${
            r.link
              ? `<a class="lt-name" style="color:${r.color}" href="${escapeAttr(r.link)}" target="_blank" rel="noopener noreferrer">${svgEscape(r.name)}</a>`
              : `<span class="lt-name" style="color:${r.color}">${svgEscape(r.name)}</span>`
          }
          ${r.uncertain ? '<span class="lt-warn">⚠ 待確認</span>' : ""}
          ${r.attrs ? `<span class="lt-attrs">${r.attrs}</span>` : ""}
        </div>
      `).join("")}
      <div class="list-tree-watermark">本圖由AI助手生成</div>
    </div>`;

  const listInner = document.getElementById("listTreeInner");
  if (listInner) {
    const vw = container.clientWidth - 24;
    const vh = container.clientHeight - 24;
    const rawW = listInner.scrollWidth;
    const rawH = listInner.scrollHeight;
    const fitScale = Math.min(vw / Math.max(rawW, 1), vh / Math.max(rawH, 1), 1);
    const shouldFitOnePage = state.chartIntent === "print_single";
    const scale = shouldFitOnePage ? Math.max(0.62, fitScale) : 1;
    listInner.style.transformOrigin = "top left";
    listInner.style.transform = `scale(${scale.toFixed(3)})`;
  }
}

// ── ECharts 視覺圖（≤20 家） ─────────────────────────────────
function renderEChart() {
  if (!window.echarts) { console.warn("ECharts 未載入"); return; }
  if (_chart) { _chart.dispose(); _chart = null; }

  const container = elements.chartContainer;
  _chart = echarts.init(container, null, { renderer: "canvas" });

  const chartRows = getChartRows();
  const treeData = buildEChartsTree(chartRows);
  if (!treeData) return;

  const option = {
    backgroundColor: "#f8fafc",
    tooltip: {
      trigger: "item",
      enterable: false,
      formatter(params) {
        const r = params.data._row;
        if (!r) return "";
        return [
          `<b>${r.canonical_name || r.chart1_name}</b>`,
          r.legal_representative  ? `法代：${r.legal_representative}` : "",
          r.registered_capital    ? `資本額：${formatCapital(r.registered_capital)}` : "",
          r.established_date      ? `成立：${r.established_date}` : "",
          r.actual_controller_share ? `持股：${r.actual_controller_share}` : "",
          r.company_status        ? `狀態：${r.company_status}` : "",
          r.node_status !== "enriched" ? `<span style="color:#f59e0b">⚠ 資料待確認</span>` : "",
        ].filter(Boolean).join("<br/>");
      },
    },
    series: [{
      type: "tree", orient: "TB",
      data: [treeData],
      top: "5%", bottom: "5%", left: "6%", right: "6%",
      symbol: "rect", symbolSize: [NODE_W, NODE_H],
      edgeShape: "polyline", layout: "orthogonal",
      roam: true, initialTreeDepth: -1,
      label: {
        show: true, position: "inside",
        verticalAlign: "middle", align: "center",
        rich: {
          name: { fontSize: 13, fontWeight: "bold", color: "#fff", lineHeight: 22, align: "center" },
          info: { fontSize: 10.5, color: "rgba(255,255,255,0.92)", lineHeight: 18, align: "center" },
        },
      },
      leaves: { label: { position: "inside", verticalAlign: "middle", align: "center" } },
      lineStyle: { color: "#94a3b8", width: 1.5, curveness: 0 },
      emphasis: { focus: "descendant" },
      animationDurationUpdate: 500,
    }],
  };

  _chart.setOption(option);

  const levels = [...new Set(chartRows.map((r) => Number(r.chart1_level) || 0))].sort();
  elements.chartLegend.innerHTML = [
    ...levels.map((l) => {
      const color = LEVEL_COLORS[Math.min(l, LEVEL_COLORS.length - 1)];
      return `<span class="legend-item"><span class="legend-dot" style="background:${color}"></span>${LEVEL_NAMES[l] || `第${l}層`}</span>`;
    }),
    `<span class="legend-item"><span class="legend-dot legend-dot-uncertain"></span>待確認</span>`,
  ].join("");

  new ResizeObserver(() => _chart && _chart.resize()).observe(container);
}

// ── 主入口 ────────────────────────────────────────────────────
function renderChart() {
  if (!state.started || !state.masterRows.length) return;

  const chartRows = getChartRows();
  const total = chartRows.length;
  const profile = getChartProfile();
  syncChartModeButtons();
  const isList = state.chartView === "list";
  const title = getChartTitle();
  if (elements.printChartTitle) elements.printChartTitle.textContent = getPrintTitle();
  const intentLabel = {
    presentation: "簡報模式",
    print_single: "列印一頁",
    print_paged: "分頁列印",
  }[state.chartIntent] || "簡報模式";
  const hybridLabel = state.hybridMode === "on"
    ? "清單柱：強制"
    : state.hybridMode === "off"
      ? "清單柱：關閉"
      : `清單柱：自動≥${state.hybridThreshold}`;
  const densityLabel = state.chartDensity === "compact" ? "精簡" : state.chartDensity === "full" ? "完整" : "標準";

  elements.chartLayoutBadge.textContent = isList
    ? `${total} 家公司 · ${intentLabel} · 條列層級 · ${getChartDepthLabel()}`
    : `${total} 家公司 · ${intentLabel} · ${hybridLabel} · 字體${state.chartFontScale}% · 內容${densityLabel} · ${profile.label} · ${state.chartDirection === "right" ? "左到右" : "上到下"} · ${getChartDepthLabel()} · ${state.chartStyle === "mono" ? "黑白正式" : "層級彩色"}${state.showGroupRoot ? " · 含集團主體" : ""}`;

  // 切換容器樣式
  elements.chartContainer.classList.toggle("chart-container-list", isList);
  elements.chartContainer.classList.toggle("chart-container-echart", false);
  elements.chartContainer.classList.toggle("chart-container-elk", !isList);
  elements.chartContainer.classList.toggle("chart-container-paged", !isList && state.chartMode === "paged");
  document.getElementById("chart")?.classList.toggle("chart-view-list", isList);
  document.getElementById("chart")?.classList.toggle("chart-view-graph", !isList);
  document.getElementById("chart")?.classList.toggle("chart-style-mono", state.chartStyle === "mono");
  document.getElementById("chart")?.classList.toggle("chart-style-color", state.chartStyle === "color");

  elements.exportPngBtn.style.display  = isList ? "none" : "";
  elements.exportHtmlBtn.style.display = "";

  if (isList) {
    if (_chart) { _chart.dispose(); _chart = null; }
    renderListTree();
  } else {
    renderElkChart();
  }
}

function exportPNG() {
  const svg = document.querySelector(".elk-svg");
  if (svg) {
    const serializer = new XMLSerializer();
    const source = serializer.serializeToString(svg);
    const blob = new Blob([source], { type: "image/svg+xml;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const img = new Image();
    img.onload = () => {
      const canvas = document.createElement("canvas");
      canvas.width = img.width * 2;
      canvas.height = img.height * 2;
      const ctx = canvas.getContext("2d");
      ctx.fillStyle = "#f8fafc";
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      ctx.scale(2, 2);
      ctx.drawImage(img, 0, 0);
      URL.revokeObjectURL(url);
      const a = document.createElement("a");
      a.href = canvas.toDataURL("image/png");
      a.download = `${state.taskName || "股權架構圖"}.png`;
      a.click();
    };
    img.src = url;
    return;
  }

  if (!_chart) return;
  const url = _chart.getDataURL({ type: "png", pixelRatio: 2, backgroundColor: "#f8fafc" });
  const a = document.createElement("a");
  a.href = url;
  a.download = `${state.taskName || "股權架構圖"}.png`;
  a.click();
}

function exportHTML() {
  const title = getChartTitle();
  const svgMarkup = document.getElementById("elkPagedChart")?.outerHTML || document.querySelector(".elk-svg")?.outerHTML;
  if (svgMarkup) {
    const profile = getChartProfile();
    const html = `<!DOCTYPE html>
<html lang="zh-Hant"><head>
<meta charset="UTF-8"><title>${svgEscape(title)}</title>
<style>
@page { size: ${state.chartMode === "a3" ? "A3" : "A4"} landscape; margin: 12mm; }
body { margin: 0; background: #f8fafc; font-family: "Noto Sans TC", "PingFang TC", sans-serif; }
.wrap { padding: 18px; }
h2 { margin: 0 0 12px; color: #1e293b; font-size: 16px; }
svg { max-width: 100%; height: auto; background: #f8fafc; }
.elk-page { break-after: page; page-break-after: always; margin-bottom: 24px; }
.elk-page:last-child { break-after: auto; page-break-after: auto; }
.elk-page-title { color: #425466; font-size: 12px; font-weight: 700; margin: 0 0 8px; }
</style></head>
<body><div class="wrap"><h2>${svgEscape(title)} — ${profile.label}</h2>${svgMarkup}</div></body></html>`;
    const blob = new Blob([html], { type: "text/html;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${title}.html`;
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 3000);
    return;
  }

  const listInner = document.getElementById("listTreeInner")?.outerHTML || "";
  if (listInner) {
    const html = `<!DOCTYPE html>
<html lang="zh-Hant"><head>
<meta charset="UTF-8"><title>${svgEscape(title)}</title>
<style>
@page { size: A4 landscape; margin: 15mm; }
body { font-family: "Noto Sans TC", "PingFang TC", sans-serif; background: #f8fafc; padding: 24px; }
h2 { font-size: 16px; margin-bottom: 16px; color: #1e293b; }
.list-tree { font-size: 12px; line-height: 1.9; }
.lt-row { display: flex; align-items: baseline; gap: 6px; white-space: nowrap; }
.lt-root { font-size: 15px; font-weight: 800; margin-bottom: 4px; }
.lt-prefix { font-family: monospace; color: #94a3b8; white-space: pre; }
.lt-name { font-weight: 600; }
.lt-warn { font-size: 11px; color: #f59e0b; font-weight: 600; }
.lt-attrs { color: #64748b; font-size: 11px; }
.lt-uncertain .lt-name { text-decoration: underline dotted #f59e0b; }
</style></head>
<body>
<h2>${svgEscape(title)} — 條列層級</h2>
${listInner}
</body></html>`;
    const blob = new Blob([html], { type: "text/html;charset=utf-8" });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement("a");
    a.href = url; a.download = `${title}.html`; a.click();
    setTimeout(() => URL.revokeObjectURL(url), 3000);
    return;
  }

  const useList = state.masterRows.length > 20;

  if (useList) {
    // 條列式：匯出含樣式的獨立 HTML
    const inner = document.getElementById("listTreeInner")?.outerHTML || "";
    const html = `<!DOCTYPE html>
<html lang="zh-Hant"><head>
<meta charset="UTF-8"><title>${title}</title>
<style>
@page { size: A4 landscape; margin: 15mm; }
body { font-family: "Noto Sans TC", "PingFang TC", sans-serif; background: #f8fafc; padding: 24px; }
h2 { font-size: 16px; margin-bottom: 16px; color: #1e293b; }
.list-tree { font-size: 12px; line-height: 1.9; }
.lt-row { display: flex; align-items: baseline; gap: 6px; white-space: nowrap; }
.lt-root { font-size: 15px; font-weight: 800; margin-bottom: 4px; }
.lt-prefix { font-family: monospace; color: #94a3b8; white-space: pre; }
.lt-name { font-weight: 600; }
.lt-warn { font-size: 11px; color: #f59e0b; font-weight: 600; }
.lt-attrs { color: #64748b; font-size: 11px; }
.lt-uncertain .lt-name { text-decoration: underline dotted #f59e0b; }
</style></head>
<body>
<h2>${title} — 股權架構圖</h2>
${inner}
</body></html>`;
    const blob = new Blob([html], { type: "text/html;charset=utf-8" });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement("a");
    a.href = url; a.download = `${title}.html`; a.click();
    setTimeout(() => URL.revokeObjectURL(url), 3000);
  } else {
    // ECharts 模式
    if (!_chart) return;
    const option = JSON.stringify(_chart.getOption());
    const html = `<!DOCTYPE html>
<html lang="zh-Hant"><head>
<meta charset="UTF-8"><title>${title}</title>
<script src="https://cdn.jsdelivr.net/npm/echarts@5.4.3/dist/echarts.min.js"><\/script>
<style>*{margin:0;padding:0}body{background:#f8fafc}#c{width:100vw;height:100vh}
.tip{position:fixed;bottom:16px;right:20px;font:13px/1.5 sans-serif;color:#64748b;
background:#ffffffcc;padding:6px 12px;border-radius:8px}</style></head>
<body><div id="c"></div><div class="tip">滾輪縮放 · 拖曳移動</div>
<script>const c=echarts.init(document.getElementById('c'));
c.setOption(${option});window.addEventListener('resize',()=>c.resize());<\/script>
</body></html>`;
    const blob = new Blob([html], { type: "text/html;charset=utf-8" });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement("a");
    a.href = url; a.download = `${title}.html`; a.click();
    setTimeout(() => URL.revokeObjectURL(url), 3000);
  }
}

function openPrintSettings() {
  document.getElementById("printSettingsModal")?.remove();
  const modal = document.createElement("div");
  modal.id = "printSettingsModal";
  modal.className = "print-drawer-backdrop";
  const currentTitle = svgEscape(getPrintTitle());
  modal.innerHTML = `
    <div class="print-settings-preview" aria-hidden="true">
      <div class="print-preview-paper ${state.printOrientation}">
        <h4>${currentTitle}</h4>
        <div id="printPreviewFrame" class="print-preview-frame">
          <div id="printPreviewViewport" class="print-preview-viewport"></div>
        </div>
        <p id="printPreviewHint" class="print-preview-hint"></p>
      </div>
    </div>
    <aside class="print-settings-panel" role="dialog" aria-modal="true" aria-labelledby="printSettingsTitle">
      <div class="section-head">
        <div>
          <p class="eyebrow">列印設定</p>
          <h3 id="printSettingsTitle">正式列印版面</h3>
        </div>
      </div>
      <label class="field print-title-field">
        <span>表頭</span>
        <input id="printTitleInput" type="text" value="${currentTitle}" placeholder="輸入列印表頭" />
      </label>
      <div class="print-orientation-grid">
        <button class="print-orientation-btn" data-print-orientation="portrait" type="button">
          <span class="print-paper portrait"></span>
          <strong>直式列印</strong>
        </button>
        <button class="print-orientation-btn" data-print-orientation="landscape" type="button">
          <span class="print-paper landscape"></span>
          <strong>橫式列印</strong>
        </button>
      </div>
      <label class="print-fit-toggle">
        <input id="printFitToPageInput" type="checkbox" ${state.printFitToPage ? "checked" : ""} />
        <span>
          <strong>自動縮到一頁</strong>
          <small>系統會依紙張方向自動縮放並置中，避免手動排版。</small>
        </span>
      </label>
      <div class="field">
        <span>快速套用</span>
        <div class="print-orientation-grid">
          <button class="ghost-btn" data-print-preset="fit" type="button">一頁優先</button>
          <button class="ghost-btn" data-print-preset="balance" type="button">平衡</button>
          <button class="ghost-btn" data-print-preset="readable" type="button">可讀優先</button>
        </div>
      </div>
      <label class="field">
        <span>縮放（%）</span>
        <input id="printScaleInput" type="number" min="70" max="130" step="5" value="${state.printScale}" />
      </label>
      <label class="field">
        <span>邊界</span>
        <select id="printMarginSelect">
          <option value="narrow" ${state.printMargin === "narrow" ? "selected" : ""}>窄</option>
          <option value="normal" ${state.printMargin === "normal" ? "selected" : ""}>標準</option>
          <option value="wide" ${state.printMargin === "wide" ? "selected" : ""}>寬</option>
        </select>
      </label>
      <label class="field">
        <span>字級</span>
        <select id="printFontSizeSelect">
          <option value="small" ${state.printFontSize === "small" ? "selected" : ""}>小</option>
          <option value="medium" ${state.printFontSize === "medium" ? "selected" : ""}>中</option>
          <option value="large" ${state.printFontSize === "large" ? "selected" : ""}>大</option>
        </select>
      </label>
      <label class="field">
        <span>間距</span>
        <select id="printSpacingSelect">
          <option value="compact" ${state.printSpacing === "compact" ? "selected" : ""}>緊湊</option>
          <option value="normal" ${state.printSpacing === "normal" ? "selected" : ""}>標準</option>
          <option value="loose" ${state.printSpacing === "loose" ? "selected" : ""}>寬鬆</option>
        </select>
      </label>
      <label class="print-fit-toggle">
        <input id="printForceOnePageInput" type="checkbox" ${state.printForceOnePage ? "checked" : ""} />
        <span>
          <strong>強制一頁</strong>
          <small>可能會變小，僅在必要時使用。</small>
        </span>
      </label>
      <div class="detail-actions">
        <button class="ghost-btn" id="printSettingsCancel" type="button">取消</button>
        <button class="primary-btn" id="printSettingsSubmit" type="button">開啟列印</button>
      </div>
    </aside>
  `;
  document.body.appendChild(modal);
  let selectedOrientation = state.printOrientation;
  const previewPaper = modal.querySelector(".print-preview-paper");
  modal.querySelectorAll("[data-print-orientation]").forEach((button) => {
    button.classList.toggle("active", button.dataset.printOrientation === state.printOrientation);
    button.addEventListener("click", () => {
      selectedOrientation = button.dataset.printOrientation;
      modal.querySelectorAll("[data-print-orientation]").forEach((item) => {
        item.classList.toggle("active", item === button);
      });
      previewPaper?.classList.toggle("portrait", selectedOrientation === "portrait");
      previewPaper?.classList.toggle("landscape", selectedOrientation !== "portrait");
      requestAnimationFrame(() => renderPrintPreview());
    });
  });
  modal.querySelector("#printTitleInput")?.addEventListener("input", (event) => {
    const title = event.target.value.trim() || getChartTitle();
    const h4 = modal.querySelector(".print-preview-paper h4");
    if (h4) h4.textContent = title;
  });
  const fitInput = modal.querySelector("#printFitToPageInput");
  const scaleInput = modal.querySelector("#printScaleInput");
  const marginSelect = modal.querySelector("#printMarginSelect");
  const fontSizeSelect = modal.querySelector("#printFontSizeSelect");
  const spacingSelect = modal.querySelector("#printSpacingSelect");
  const forceOnePageInput = modal.querySelector("#printForceOnePageInput");
  const previewViewport = modal.querySelector("#printPreviewViewport");
  const previewHint = modal.querySelector("#printPreviewHint");
  const renderPrintPreview = () => {
    if (!previewViewport) return;
    const title = (modal.querySelector("#printTitleInput")?.value || "").trim() || getChartTitle();
    const h4 = modal.querySelector(".print-preview-paper h4");
    if (h4) h4.textContent = title;
    const sourceSvg = document.querySelector("#chartContainer .elk-svg");
    const sourceList = document.querySelector("#chartContainer #listTreeInner");
    if (sourceSvg) {
      previewViewport.innerHTML = sourceSvg.outerHTML;
      if (previewHint) previewHint.textContent = "";
      const previewSvg = previewViewport.querySelector("svg");
      if (!previewSvg) return;
      previewSvg.removeAttribute("id");
      const rawWidth = Number(previewSvg.getAttribute("width")) || 1200;
      const rawHeight = Number(previewSvg.getAttribute("height")) || 900;
      const vw = previewViewport.clientWidth || 700;
      const vh = previewViewport.clientHeight || 440;
      const fitToPage = Boolean(fitInput?.checked || forceOnePageInput?.checked);
      const manualScale = clampNumber(Number(scaleInput?.value || 100), 70, 130) / 100;
      const baseScale = Math.min(vw / rawWidth, vh / rawHeight, 1);
      const scale = fitToPage ? baseScale * manualScale : Math.min(1.25, baseScale * manualScale);
      previewSvg.style.width = `${Math.max(220, rawWidth * scale)}px`;
      previewSvg.style.height = "auto";
      previewSvg.style.display = "block";
      previewSvg.style.margin = "0 auto";
      return;
    }
    if (sourceList) {
      previewViewport.innerHTML = `<div id="printPreviewList" class="print-preview-list">${sourceList.innerHTML}<div class="list-tree-watermark">本圖由AI助手生成</div></div>`;
      const wrapper = previewViewport.querySelector("#printPreviewList");
      if (wrapper) {
        const fitToPage = Boolean(fitInput?.checked || forceOnePageInput?.checked);
        const manualScale = clampNumber(Number(scaleInput?.value || 100), 70, 130) / 100;
        const vw = previewViewport.clientWidth || 700;
        const vh = previewViewport.clientHeight || 440;
        const rawW = wrapper.scrollWidth || vw;
        const rawH = wrapper.scrollHeight || vh;
        const baseScale = Math.min(vw / rawW, vh / rawH, 1);
        const scale = fitToPage ? Math.max(0.4, baseScale * manualScale) : Math.min(1.2, baseScale * manualScale);
        wrapper.style.transform = `scale(${scale.toFixed(3)})`;
        wrapper.style.transformOrigin = "top left";
        wrapper.style.width = `${rawW}px`;
        wrapper.style.height = `${rawH}px`;
        const projectedH = rawH * scale;
        const isOnePage = projectedH <= vh + 1;
        if (previewHint) previewHint.textContent = isOnePage ? "預覽：可放入一頁" : "預覽：超過一頁，請調整縮放或字距";
      }
      return;
    }
    previewViewport.innerHTML = `<div class="print-preview-empty">目前沒有可預覽圖面</div>`;
    if (previewHint) previewHint.textContent = "";
  };
  const applyPrintPreset = (preset) => {
    if (!fitInput || !scaleInput || !marginSelect || !fontSizeSelect || !spacingSelect || !forceOnePageInput) return;
    if (preset === "fit") {
      fitInput.checked = true;
      scaleInput.value = "90";
      marginSelect.value = "narrow";
      fontSizeSelect.value = "small";
      spacingSelect.value = "compact";
      forceOnePageInput.checked = true;
      renderPrintPreview();
      return;
    }
    if (preset === "readable") {
      fitInput.checked = false;
      scaleInput.value = "110";
      marginSelect.value = "normal";
      fontSizeSelect.value = "large";
      spacingSelect.value = "loose";
      forceOnePageInput.checked = false;
      renderPrintPreview();
      return;
    }
    fitInput.checked = true;
    scaleInput.value = "100";
    marginSelect.value = "normal";
    fontSizeSelect.value = "medium";
    spacingSelect.value = "normal";
    forceOnePageInput.checked = false;
    renderPrintPreview();
  };
  modal.querySelectorAll("[data-print-preset]").forEach((button) => {
    button.addEventListener("click", () => applyPrintPreset(button.dataset.printPreset));
  });
  [fitInput, scaleInput, forceOnePageInput, marginSelect, fontSizeSelect, spacingSelect].forEach((input) => {
    input?.addEventListener("input", renderPrintPreview);
    input?.addEventListener("change", renderPrintPreview);
  });
  modal.querySelector("#printSettingsSubmit")?.addEventListener("click", async () => {
    const titleInput = modal.querySelector("#printTitleInput");
    state.printTitle = titleInput?.value.trim() || getChartTitle();
    state.printFitToPage = Boolean(fitInput?.checked);
    state.printScale = clampNumber(Number(scaleInput?.value || 100), 70, 130);
    state.printMargin = ["narrow", "normal", "wide"].includes(marginSelect?.value) ? marginSelect.value : "normal";
    state.printFontSize = ["small", "medium", "large"].includes(fontSizeSelect?.value) ? fontSizeSelect.value : "medium";
    state.printSpacing = ["compact", "normal", "loose"].includes(spacingSelect?.value) ? spacingSelect.value : "normal";
    state.printForceOnePage = Boolean(forceOnePageInput?.checked);
    await savePrintSettings(selectedOrientation).catch((error) => console.error("print settings save failed", error));
    modal.remove();
    printChart(selectedOrientation);
  });
  modal.querySelector("#printSettingsCancel")?.addEventListener("click", () => modal.remove());
  modal.addEventListener("click", (event) => {
    if (event.target === modal) modal.remove();
  });
  requestAnimationFrame(renderPrintPreview);
}

async function savePrintSettings(orientation = state.printOrientation) {
  if (!state.taskId) return;
  const payload = await apiPost(`/api/tasks/${state.taskId}/chart-print-settings`, {
    title: state.printTitle || getChartTitle(),
    orientation,
    fit_to_page: state.printFitToPage,
    scale: state.printScale,
    margin: state.printMargin,
    font_size: state.printFontSize,
    spacing: state.printSpacing,
    force_one_page: state.printForceOnePage,
  });
  applyTaskRefresh(payload);
}

function printChart(orientation = state.printOrientation) {
  state.printOrientation = orientation === "portrait" ? "portrait" : "landscape";
  if (elements.printChartTitle) elements.printChartTitle.textContent = getPrintTitle();
  document.body.classList.toggle("print-portrait", state.printOrientation === "portrait");
  document.body.classList.toggle("print-landscape", state.printOrientation !== "portrait");
  preparePrintLayout();
  document.getElementById("dynamicPrintPageStyle")?.remove();
  const style = document.createElement("style");
  style.id = "dynamicPrintPageStyle";
  const marginMap = {
    narrow: "8mm 10mm",
    normal: "12mm 15mm",
    wide: "16mm 20mm",
  };
  const margin = marginMap[state.printMargin] || marginMap.normal;
  const fontScale = {
    small: 0.94,
    medium: 1,
    large: 1.08,
  }[state.printFontSize] || 1;
  const chartFontScale = clampNumber(Number(state.chartFontScale || 100), 85, 120) / 100;
  const spacingScale = {
    compact: 0.92,
    normal: 1,
    loose: 1.1,
  }[state.printSpacing] || 1;
  style.textContent = `@media print {
    @page { size: A4 ${state.printOrientation}; margin: ${margin}; }
    #chart { --print-font-scale: ${(fontScale * chartFontScale).toFixed(4)}; --print-spacing-scale: ${spacingScale}; }
  }`;
  document.head.appendChild(style);
  window.print();
}

function preparePrintLayout() {
  const container = elements.chartContainer;
  if (!container) return;
  const marginWidthByProfile = { narrow: 6, normal: 0, wide: -10 };
  const marginHeightByProfile = { narrow: 8, normal: 0, wide: -12 };
  const maxWidthMmBase = state.printOrientation === "portrait" ? 180 : 267;
  const maxHeightMmBase = state.printOrientation === "portrait" ? 235 : 148;
  const maxWidthMm = maxWidthMmBase + (marginWidthByProfile[state.printMargin] || 0);
  const maxHeightMm = maxHeightMmBase + (marginHeightByProfile[state.printMargin] || 0);
  const manualScale = clampNumber(Number(state.printScale || 100), 70, 130) / 100;
  container.querySelectorAll(".elk-svg").forEach((svg) => {
    const width = Number(svg.getAttribute("width")) || svg.viewBox?.baseVal?.width || 1;
    const height = Number(svg.getAttribute("height")) || svg.viewBox?.baseVal?.height || 1;
    const fitEnabled = state.printForceOnePage || state.printFitToPage;
    const autoScale = fitEnabled ? Math.min(maxWidthMm / width, maxHeightMm / height, 1) : Math.min(maxWidthMm / width, 1);
    const scale = Math.min(1.4, autoScale * manualScale);
    const printWidth = Math.max(60, Math.min(maxWidthMm, width * scale));
    svg.style.setProperty("--print-svg-width", `${printWidth.toFixed(2)}mm`);
    svg.style.setProperty("--print-svg-height", "auto");
  });
}

function bindEvents() {
  elements.navButtons.forEach((button) => {
    button.addEventListener("click", () => setView(button.dataset.view));
  });
  elements.chart1Input.addEventListener("change", (event) => {
    state.chart1File = event.target.files[0];
    setPreview(state.chart1File, elements.chart1Meta, elements.chart1Preview, document.getElementById("dz1"));
    updateImageCheck("chart1", state.chart1File);
    enableStartIfReady();
  });
  elements.chart2Input.addEventListener("change", (event) => {
    state.chart2File = event.target.files[0];
    setPreview(state.chart2File, elements.chart2Meta, elements.chart2Preview, document.getElementById("dz2"));
    updateImageCheck("chart2", state.chart2File);
    enableStartIfReady();
  });
  elements.ocrTestInput?.addEventListener("change", (event) => {
    state.ocrTestFile = event.target.files?.[0] || null;
    if (state.ocrTestFile) {
      setPreview(
        state.ocrTestFile,
        elements.ocrTestMeta,
        elements.ocrTestPreview,
        document.getElementById("ocrTestDropzone"),
      );
    }
    if (elements.ocrTestResult) {
      elements.ocrTestResult.className = "ocr-test-empty";
      elements.ocrTestResult.textContent = "已選擇圖片，可以開始 OCR 測試。";
    }
    enableOcrTestIfReady();
  });
  elements.ocrProviderSelect?.addEventListener("change", () => {
    if (elements.ocrTestResult && !state.ocrTesting) {
      elements.ocrTestResult.className = "ocr-test-empty";
      elements.ocrTestResult.textContent = "已切換 Provider，可以重新開始 OCR 測試。";
    }
  });
  elements.ocrPromptProfileSelect?.addEventListener("change", () => {
    if (elements.ocrTestResult && !state.ocrTesting) {
      elements.ocrTestResult.className = "ocr-test-empty";
      elements.ocrTestResult.textContent = "已切換 Prompt 版本，可以重新開始 OCR 測試。";
    }
  });
  elements.ocrTestBtn?.addEventListener("click", runOcrTest);
  elements.adminUnlockBtn?.addEventListener("click", unlockAdminTest);
  elements.adminPasswordInput?.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      unlockAdminTest();
    }
  });
  elements.ocrRefreshHistoryBtn?.addEventListener("click", () => {
    loadOcrTestHistory().catch((error) => renderOcrTestError(error.message));
  });
  elements.taskNameInput.addEventListener("input", (event) => {
    state.taskName = event.target.value.trim();
    if (state.taskName) document.getElementById("uploadError")?.remove();
  });
  elements.uploadModeButtons.forEach((btn) => {
    btn.addEventListener("click", () => setUploadMode(btn.dataset.uploadMode || "ocr"));
  });
  elements.startAnalysisBtn.addEventListener("click", async () => {
    const originalText = elements.startAnalysisBtn.textContent;
    document.getElementById("uploadError")?.remove();
    if (!elements.taskNameInput.value.trim()) {
      const errDiv = document.createElement("div");
      errDiv.id = "uploadError";
      errDiv.className = "upload-error-msg";
      errDiv.innerHTML = `<strong>請先填寫集團名稱</strong><br><small>系統會用它作為結果主表與股權架構圖標題。</small>`;
      elements.startAnalysisBtn.closest(".cta-row")?.after(errDiv);
      elements.taskNameInput.focus();
      return;
    }
    try {
      state.loading = true;
      enableStartIfReady();
      await createTaskFromUpload((statusMsg) => {
        elements.startAnalysisBtn.textContent = statusMsg;
      });
    } catch (error) {
      console.error(error);
      const errDiv = document.createElement("div");
      errDiv.id = "uploadError";
      errDiv.className = "upload-error-msg";
      errDiv.innerHTML = `<strong>分析失敗</strong>：${error.message}<br><small>請確認圖片清晰度，或稍後再試。</small>`;
      elements.startAnalysisBtn.closest(".cta-row")?.after(errDiv);
    } finally {
      state.loading = false;
      elements.startAnalysisBtn.textContent = originalText;
      enableStartIfReady();
    }
  });
  elements.createManualTaskBtn?.addEventListener("click", async () => {
    const originalText = elements.createManualTaskBtn.textContent;
    document.getElementById("uploadError")?.remove();
    try {
      state.loading = true;
      elements.createManualTaskBtn.textContent = "建立中…";
      await createManualTask();
    } catch (error) {
      const errDiv = document.createElement("div");
      errDiv.id = "uploadError";
      errDiv.className = "upload-error-msg";
      errDiv.innerHTML = `<strong>建立失敗</strong>：${error.message}`;
      elements.manualCreatePanel?.appendChild(errDiv);
    } finally {
      state.loading = false;
      elements.createManualTaskBtn.textContent = originalText;
      enableStartIfReady();
    }
  });
  elements.searchInput?.addEventListener("input", renderResults);
  elements.statusFilter?.addEventListener("change", renderResults);
  elements.taskSearchBtn?.addEventListener("click", () => { loadTaskCenter().catch((e) => console.error(e)); });
  elements.taskSearchInput?.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      loadTaskCenter().catch((e) => console.error(e));
    }
  });
  elements.taskStatusSelect?.addEventListener("change", () => { loadTaskCenter().catch((e) => console.error(e)); });
  elements.batchApplyBtn?.addEventListener("click", () => {
    runBatchUpdate().catch((error) => alert(`批次更新失敗：${error.message}`));
  });
  elements.batchDeleteBtn?.addEventListener("click", () => {
    runBatchDelete().catch((error) => alert(`批次刪除失敗：${error.message}`));
  });
  elements.openShareholderModalBtn?.addEventListener("click", openShareholderModal);
  elements.openExternalEntityModalBtn?.addEventListener("click", openExternalEntityModal);
  elements.undoBtn?.addEventListener("click", () => {
    undoTaskEdit().catch((error) => alert(`回復失敗：${error.message}`));
  });
  elements.redoBtn?.addEventListener("click", () => {
    redoTaskEdit().catch((error) => alert(`重做失敗：${error.message}`));
  });
  elements.saveDraftBtn?.addEventListener("click", () => {
    saveDraftSnapshot(false).catch((error) => console.error("draft save failed", error));
  });
  elements.restoreDraftBtn?.addEventListener("click", () => {
    restoreDraftSnapshot().catch((error) => alert(`還原草稿失敗：${error.message}`));
  });
  elements.addCompanyBtn?.addEventListener("click", () => setAddCompanyPanel(true));
  elements.bulkAddBtn?.addEventListener("click", () => setBulkAddPanel(true));
  elements.cancelAddCompanyBtn?.addEventListener("click", () => setAddCompanyPanel(false));
  elements.saveAddCompanyBtn?.addEventListener("click", addCompanyToResults);
  elements.cancelBulkAddBtn?.addEventListener("click", () => setBulkAddPanel(false));
  elements.saveBulkAddBtn?.addEventListener("click", () => {
    bulkAddCompaniesToResults().catch((error) => alert(`批量新增失敗：${error.message}`));
  });
  elements.addCompanyName?.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      addCompanyToResults();
    }
  });
  elements.addCompanyShare?.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      addCompanyToResults();
    }
  });
  elements.chartViewButtons.forEach((button) => {
    button.addEventListener("click", () => {
      state.chartView = button.dataset.chartView || "graph";
      resetChartViewport();
      renderChart();
    });
  });
  elements.chartIntentButtons.forEach((button) => {
    button.addEventListener("click", () => {
      applyChartIntent(button.dataset.chartIntent || "presentation", { fromUser: true });
      state.selectedBranchId = "__all__";
      resetChartViewport();
      renderChart();
    });
  });
  elements.chartDirectionButtons.forEach((button) => {
    button.addEventListener("click", () => {
      state.chartDirection = button.dataset.chartDirection || "down";
      if (state.chartIntent !== "presentation") state.chartIntent = "presentation";
      resetChartViewport();
      renderChart();
    });
  });
  elements.chartStyleButtons.forEach((button) => {
    button.addEventListener("click", () => {
      state.chartStyle = button.dataset.chartStyle || "mono";
      renderChart();
    });
  });
  elements.chartModeButtons.forEach((button) => {
    button.addEventListener("click", () => {
      state.chartMode = button.dataset.chartMode || "a4";
      if (state.chartIntent !== "presentation") state.chartIntent = "presentation";
      state.selectedBranchId = "__all__";
      resetChartViewport();
      renderChart();
    });
  });
  elements.chartDepthButtons.forEach((button) => {
    button.addEventListener("click", () => {
      state.chartDepth = button.dataset.chartDepth || "all";
      if (state.chartIntent !== "presentation") state.chartIntent = "presentation";
      state.selectedBranchId = "__all__";
      resetChartViewport();
      renderChart();
    });
  });
  elements.chartBranchSelect?.addEventListener("change", (event) => {
    state.selectedBranchId = event.target.value || "__all__";
    resetChartViewport();
    renderChart();
  });
  elements.chartZoomButtons.forEach((button) => {
    button.addEventListener("click", () => {
      const action = button.dataset.chartZoom;
      if (action === "in") setChartZoom(state.chartScale + 0.12);
      if (action === "out") setChartZoom(state.chartScale - 0.12);
      if (action === "fit") fitChartToViewport();
      if (action === "reset") resetChartViewport();
    });
  });
  elements.chartShowRootToggle?.addEventListener("change", (event) => {
    state.showGroupRoot = event.target.checked;
    state.selectedBranchId = "__all__";
    resetChartViewport();
    renderChart();
  });
  elements.hybridModeSelect?.addEventListener("change", (event) => {
    state.hybridMode = ["auto", "off", "on"].includes(event.target.value) ? event.target.value : "auto";
    resetChartViewport();
    renderChart();
  });
  elements.hybridThresholdSelect?.addEventListener("change", (event) => {
    const n = Number(event.target.value || HYBRID_COLUMN_THRESHOLD);
    state.hybridThreshold = Math.max(2, Math.min(20, Number.isFinite(n) ? n : HYBRID_COLUMN_THRESHOLD));
    resetChartViewport();
    renderChart();
  });
  elements.chartFontScaleSelect?.addEventListener("change", (event) => {
    const n = Number(event.target.value || 100);
    state.chartFontScale = Math.max(85, Math.min(120, Number.isFinite(n) ? n : 100));
    renderChart();
  });
  elements.chartDensitySelect?.addEventListener("change", (event) => {
    state.chartDensity = ["compact", "standard", "full"].includes(event.target.value) ? event.target.value : "full";
    renderChart();
  });
  elements.toggleToolbarBtn?.addEventListener("click", () => {
    state.toolbarCollapsed = !state.toolbarCollapsed;
    applyWorkspaceModeUI();
    saveChartViewPrefs();
  });
  elements.reviewConfirmAllBtn?.addEventListener("click", () => {
    confirmAllReviewRows().catch((err) => alert(`全部確認失敗：${err.message}`));
  });
  elements.exportBtn.addEventListener("click", exportWorkbook);
  elements.exportPngBtn.addEventListener("click", exportPNG);
  elements.exportHtmlBtn.addEventListener("click", exportHTML);
  elements.printChartBtn.addEventListener("click", openPrintSettings);
}

applyChartIntent(state.chartIntent);
bindEvents();
loadChartViewPrefs();
applyWorkspaceModeUI();
applyUploadModeUI();
updateTaskBadge();
setAdminUnlocked(false);
renderActivityPanel();
syncActivityPanelVisibility("upload");
document.body.classList.add("print-landscape");
updateUndoRedoButtons();

window.addEventListener("beforeunload", (event) => {
  if (!state.hasUnsavedEdits) return;
  event.preventDefault();
  event.returnValue = "";
});

// 頁面載入時靜默 ping 後端，提前喚醒 Railway（冷啟動可能需 10–30 秒）
fetch(API_BASE + "/api/health").catch(() => {});
