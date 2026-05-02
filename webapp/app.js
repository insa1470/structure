const state = {
  taskId: "",
  taskName: "",
  chart1File: null,
  chart2File: null,
  started: false,
  loading: false,
  masterRows: [],
  reviewRows: [],
  candidateRows: [],
  reviewDecisions: {},
  candidateDecisions: {},
  selectedReviewIndex: 0,
  selectedCandidateIndex: 0,
};

const elements = {
  pageTitle: document.getElementById("pageTitle"),
  navButtons: [...document.querySelectorAll(".nav-btn")],
  views: [...document.querySelectorAll(".view")],
  chart1Input: document.getElementById("chart1Input"),
  chart2Input: document.getElementById("chart2Input"),
  chart1Meta: document.getElementById("chart1Meta"),
  chart2Meta: document.getElementById("chart2Meta"),
  chart1Preview: document.getElementById("chart1Preview"),
  chart2Preview: document.getElementById("chart2Preview"),
  taskNameInput: document.getElementById("taskNameInput"),
  startAnalysisBtn: document.getElementById("startAnalysisBtn"),
  loadDemoBtn: document.getElementById("loadDemoBtn"),
  exportBtn: document.getElementById("exportBtn"),
  modeBadge: document.getElementById("modeBadge"),
  metricsGrid: document.getElementById("metricsGrid"),
  overviewWarnings: document.getElementById("overviewWarnings"),
  reviewListTitle: document.getElementById("reviewListTitle"),
  reviewList: document.getElementById("reviewList"),
  reviewDetail: document.getElementById("reviewDetail"),
  candidateListTitle: document.getElementById("candidateListTitle"),
  candidateList: document.getElementById("candidateList"),
  candidateDetail: document.getElementById("candidateDetail"),
  searchInput: document.getElementById("searchInput"),
  statusFilter: document.getElementById("statusFilter"),
  resultTableTitle: document.getElementById("resultTableTitle"),
  resultTableBody: document.getElementById("resultTableBody"),
  chartContainer: document.getElementById("chartContainer"),
  chartLayoutBadge: document.getElementById("chartLayoutBadge"),
  chartLegend: document.getElementById("chartLegend"),
  exportPngBtn: document.getElementById("exportPngBtn"),
  exportHtmlBtn: document.getElementById("exportHtmlBtn"),
  printChartBtn: document.getElementById("printChartBtn"),
};

const pageTitles = {
  upload: "建立新任務",
  overview: "總覽",
  review: "待確認",
  candidates: "圖二新增候選",
  results: "結果主表",
  chart: "股權架構圖",
};

const API_BASE = (window.API_BASE || "").replace(/\/$/, "");

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
  elements.modeBadge.textContent = state.taskId ? `任務 ${state.taskId}` : "尚未建立任務";
  const copy = document.getElementById("statusCopy");
  if (!copy) return;
  if (state.taskId) {
    copy.textContent = "任務已載入。可從左側導覽切換各審核階段。";
  } else {
    copy.textContent = "上傳兩張圖，系統會先整理，再只把少數不確定項目交給使用者決定。";
  }
}

function setView(viewName) {
  elements.navButtons.forEach((button) => {
    button.classList.toggle("active", button.dataset.view === viewName);
  });
  elements.views.forEach((view) => {
    view.classList.toggle("active", view.id === viewName);
  });
  elements.pageTitle.textContent = pageTitles[viewName];
  if (viewName === "chart") setTimeout(renderChart, 50); // 等 DOM 顯示後再渲染
}

function enableStartIfReady() {
  elements.startAnalysisBtn.disabled = !(state.chart1File && state.chart2File) || state.loading;
}

function setPreview(file, metaEl, imgEl, dzEl) {
  if (!file) return;
  metaEl.textContent = `${file.name} · ${(file.size / 1024 / 1024).toFixed(2)} MB`;
  const objectUrl = URL.createObjectURL(file);
  imgEl.src = objectUrl;
  if (dzEl) dzEl.classList.add("has-file");
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
  // 移除舊橫幅
  document.getElementById("analysisBanner")?.remove();

  const isDemo = !task.source_files?.chart1 || task.source_files?.chart1 === "demo_chart1.png";
  const warning = task.analysis_warning;
  const mode = task.analysis_mode || "unknown";

  let msg = "", type = "";
  if (isDemo) {
    msg = "目前顯示的是示範資料，並非您上傳的圖片。";
    type = "banner-info";
  } else if (warning) {
    msg = `⚠ AI 辨識未成功，目前顯示示範資料，非本次上傳內容。<br><small>${warning}</small>`;
    type = "banner-warn";
  } else if (mode === "qwen_vl") {
    msg = `✓ AI 辨識完成（Qwen-VL），任務 ID：${task.id}`;
    type = "banner-ok";
  }

  if (!msg) return;
  const banner = document.createElement("div");
  banner.id = "analysisBanner";
  banner.className = `analysis-banner ${type}`;
  banner.innerHTML = msg;
  document.querySelector(".main")?.prepend(banner);
}

function hydrateTask(task) {
  state.taskId = task.id;
  state.taskName = task.name;
  state.masterRows = task.master_rows || [];
  state.reviewRows = (task.review_rows || []).filter((row) => row.issue_type !== "chart2_only");
  state.candidateRows = task.candidate_rows || [];
  state.reviewDecisions = task.review_decisions || {};
  state.candidateDecisions = task.candidate_decisions || {};
  state.selectedReviewIndex = 0;
  state.selectedCandidateIndex = 0;
  state.started = true;
  updateTaskBadge();
  showAnalysisBanner(task);
  renderOverview(task.summary || {});
  renderReviewList();
  renderReviewDetail();
  renderCandidateList();
  renderCandidateDetail();
  renderResults();
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

  const warnings = [
    `${pending} 筆資料仍需要人工確認，請先處理這一區。`,
    `${candidates} 筆圖二新增候選尚未決定是否加入主表。`,
    "第二階段股權架構圖會依賴這裡的最終審核結果，所以名稱與上層公司要盡量確認乾淨。",
  ];
  elements.overviewWarnings.innerHTML = warnings.map((text) => `<li>${text}</li>`).join("");
}

function renderReviewList() {
  elements.reviewListTitle.textContent = `${state.reviewRows.length} 筆待確認`;
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

function renderReviewDetail() {
  const row = state.reviewRows[state.selectedReviewIndex];
  if (!row) return;
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
    state.reviewDecisions = response.review_decisions || {};
    renderReviewList();
    renderResults();
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
      return `
        <article class="review-item ${index === state.selectedCandidateIndex ? "active" : ""}" data-candidate-index="${index}">
          <h4 class="candidate-title">${row.chart2_name}</h4>
          <div class="pill-row">
            <span class="pill slate">${row.subsidiary_level_label || "未標級別"}</span>
            <span class="pill info">${row.company_status || "未標狀態"}</span>
            ${decision?.decision ? `<span class="pill warning">已填：${decision.decision}</span>` : ""}
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
    state.candidateDecisions = response.candidate_decisions || {};
    renderCandidateList();
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

async function reparentNode(nodeId, newParentId) {
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
  return roots;
}

function flattenTree(nodes, depth = 0, result = []) {
  nodes.forEach((node) => {
    result.push({ ...node, _depth: depth });
    if (node.children?.length) flattenTree(node.children, depth + 1, result);
  });
  return result;
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
    const finish = async (save) => {
      nameSpan.contentEditable = "false";
      const newVal = nameSpan.textContent.trim();
      if (save && newVal && newVal !== curName) {
        row.canonical_name = newVal;
        try {
          const result = await apiPost(`/api/tasks/${state.taskId}/update-row`, {
            node_id: row.node_id, canonical_name: newVal,
          });
          state.masterRows = result.master_rows || state.masterRows;
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
      if (ev.key === "Enter")  { ev.preventDefault(); nameSpan.blur(); }
      if (ev.key === "Escape") { finish(false); }
    }, { once: true });
  });
}

function unformatCapital(str) {
  return (str || "").replace(/,/g, "");
}

// 連動更新欄位（修改一個，同值的全部跟著改）
const CASCADE_FIELDS = new Set(["legal_representative"]);

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
      cell.textContent = display || "—";

      if (newRaw === rawOriginal) return; // 沒有改變

      try {
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
        state.masterRows = result.master_rows || state.masterRows;
        // 連動時重新渲染整張表
        if (CASCADE_FIELDS.has(field)) renderResults();
      } catch (e) {
        console.error("儲存失敗", e);
        cell.textContent = displayValue || "—";
      }
    };

    input.addEventListener("blur", save);
    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") { e.preventDefault(); input.blur(); }
      if (e.key === "Escape") { cell.textContent = displayValue || "—"; }
    });
  });
}

function renderResults() {
  const query  = elements.searchInput.value.trim();
  const filter = elements.statusFilter.value;

  const filteredRows = state.masterRows.filter((row) => {
    const inFilter = filter === "all" || row.node_status === filter;
    const haystack = [row.canonical_name, row.chart1_name, row.chart1_parent_name, row.legal_representative]
      .join(" ").toLowerCase();
    return inFilter && (!query || haystack.includes(query.toLowerCase()));
  });

  const rows = (query || filter !== "all")
    ? filteredRows
    : flattenTree(buildTree(state.masterRows));

  elements.resultTableTitle.textContent = `${filteredRows.length} 家公司`;

  // ── 動態層級欄數 ────────────────────────────────────────────
  const maxLevel = Math.max(0, ...state.masterRows.map((r) => Number(r.chart1_level) || 0));
  const LEVEL_HEADERS = { 0: "集團主體", 1: "一級子公司", 2: "二級子公司", 3: "三級子公司", 4: "四級子公司" };

  // 更新表頭
  const theadTr = document.querySelector("#results table thead tr");
  let headHtml = `<th class="del-col"></th>`;
  for (let lv = 0; lv <= maxLevel; lv++) {
    headHtml += `<th class="level-col">${LEVEL_HEADERS[lv] || `${lv}級子公司`}</th>`;
  }
  headHtml += `<th class="editable-col">法定代表人</th>
    <th class="editable-col">資本額</th>
    <th class="editable-col">成立日期</th>
    <th class="editable-col">持股%</th>
    <th class="editable-col">狀態</th>
    <th class="status-col">系統</th>`;
  theadTr.innerHTML = headHtml;

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
    const statusClass = row.node_status === "enriched" ? "status-enriched"
      : row.node_status === "review_match" ? "status-review" : "status-slate";

    const tr = document.createElement("tr");
    tr.className = statusClass;
    tr.dataset.nodeId = row.node_id;
    const bg = groupBgMap[row.node_id];
    if (bg) tr.style.backgroundColor = bg;

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
        const result = await apiPost(`/api/tasks/${state.taskId}/delete-row`, { node_id: row.node_id });
        state.masterRows = result.master_rows || state.masterRows;
        renderResults();
      } catch (err) { console.error("刪除失敗", err); }
    });
    delTd.appendChild(delBtn);
    tr.appendChild(delTd);

    // ── 拖曳事件（整行） ────────────────────────────────────
    tr.draggable = true;
    tr.addEventListener("dragstart", (e) => {
      _dragNodeId = row.node_id;
      e.dataTransfer.effectAllowed = "move";
      e.dataTransfer.setData("text/plain", row.node_id);
      const ghost = document.createElement("div");
      ghost.textContent = row.canonical_name || row.chart1_name || "";
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
    });
    tr.addEventListener("dragover", (e) => {
      if (!_dragNodeId || _dragNodeId === row.node_id) return;
      if (isAncestor(_dragNodeId, row.node_id)) return;
      e.preventDefault();
      e.dataTransfer.dropEffect = "move";
      document.querySelectorAll(".drag-target").forEach((el) => el.classList.remove("drag-target"));
      tr.classList.add("drag-target");
      let tip = document.getElementById("drag-tooltip");
      if (!tip) { tip = document.createElement("div"); tip.id = "drag-tooltip"; document.body.appendChild(tip); }
      tip.textContent = `↳ 放入「${row.canonical_name || row.chart1_name}」底下`;
      tip.style.left = e.clientX + "px";
      tip.style.top  = e.clientY + "px";
    });
    tr.addEventListener("dragleave", (e) => { if (!tr.contains(e.relatedTarget)) tr.classList.remove("drag-target"); });
    tr.addEventListener("drop", async (e) => {
      e.preventDefault();
      tr.classList.remove("drag-target");
      document.getElementById("drag-tooltip")?.remove();
      const draggedId = e.dataTransfer.getData("text/plain");
      if (!draggedId || draggedId === row.node_id || isAncestor(draggedId, row.node_id)) return;
      await reparentNode(draggedId, row.node_id);
      document.querySelector(`tr[data-node-id="${draggedId}"]`)?.scrollIntoView({ behavior: "smooth", block: "center" });
    });

    // ── 層級欄：公司名在對應欄，其餘空白 ────────────────────
    const curName = row.canonical_name || row.chart1_name || "";
    for (let lv = 0; lv <= maxLevel; lv++) {
      const td = document.createElement("td");
      if (lv === level) {
        td.className = "tree-name-cell";
        td.innerHTML = `<span class="drag-handle" title="拖曳調整層級">⠿</span><span class="company-name editable-name" title="點擊編輯名稱">${curName}</span>`;
        attachNameEdit(td, row);
      } else {
        td.className = "level-empty-td";
      }
      tr.appendChild(td);
    }

    // 資料欄
    [
      { key: "legal_representative",    display: row.legal_representative },
      { key: "registered_capital",      display: formatCapital(row.registered_capital) },
      { key: "established_date",        display: row.established_date },
      { key: "actual_controller_share", display: row.actual_controller_share },
      { key: "company_status",          display: row.company_status },
    ].forEach(({ key, display }) => {
      const td = document.createElement("td");
      td.textContent = display || "—";
      td.className = "editable-cell";
      makeEditable(td, row, key, display);
      tr.appendChild(td);
    });

    // 系統狀態欄
    const statusTd = document.createElement("td");
    statusTd.className = "status-col";
    statusTd.textContent = statusText(row);
    tr.appendChild(statusTd);

    elements.resultTableBody.appendChild(tr);
  });
}

async function pollTask(taskId, onTick) {
  const MAX_MS = 12 * 60 * 1000; // 12 分鐘上限
  const INTERVAL = 4000;          // 每 4 秒查一次
  const start = Date.now();
  let tick = 0;
  while (Date.now() - start < MAX_MS) {
    await new Promise((r) => setTimeout(r, INTERVAL));
    tick++;
    if (onTick) onTick(tick);
    const task = await apiGet(`/api/tasks/${taskId}`);
    if (task.status === "ready") return task;
    if (task.status === "error") throw new Error(task.error || "AI 辨識失敗，請重試");
    // status === "processing" → 繼續等
  }
  throw new Error("分析逾時（超過 12 分鐘），請重試或裁切圖片後再上傳");
}

// ── 工作區渲染 ────────────────────────────────────────────────
// phase: "idle" | "uploading" | "processing" | "ready" | "error"
function renderWorkspace(phase, opts = {}) {
  const el = document.getElementById("workspaceContent");
  if (!el) return;

  // step 設定
  const steps = [
    {
      key: "upload",
      title: "上傳圖片",
      detail: {
        idle: "等待開始",
        uploading: "上傳圖一與圖二至伺服器…",
        processing: `圖一：${state.chart1File?.name || "—"} · 圖二：${state.chart2File?.name || "—"}`,
        ready: `圖一：${state.chart1File?.name || "—"} · 圖二：${state.chart2File?.name || "—"}`,
        error: `圖一：${state.chart1File?.name || "—"} · 圖二：${state.chart2File?.name || "—"}`,
      },
    },
    {
      key: "analyze",
      title: "AI 辨識圖一（結構骨架）",
      detail: {
        idle: "等待上傳",
        uploading: "等待上傳完成",
        processing: opts.msg || "Qwen-VL 辨識中…",
        ready: `辨識完成`,
        error: "辨識未完成",
      },
    },
    {
      key: "enrich",
      title: "AI 辨識圖二（補充資訊）",
      detail: {
        idle: "等待上傳",
        uploading: "等待辨識圖一",
        processing: opts.msg || "Qwen-VL 辨識中…",
        ready: "資訊補充完成",
        error: "辨識未完成",
      },
    },
  ];

  // 每個 step 的狀態
  function stepState(key) {
    if (phase === "idle") return "pending";
    if (phase === "uploading") return key === "upload" ? "active" : "pending";
    if (phase === "processing") {
      if (key === "upload") return "done";
      if (key === "analyze") return "active";
      return "pending";
    }
    if (phase === "ready") return "done";
    if (phase === "error") {
      if (key === "upload") return "done";
      return "error";
    }
    return "pending";
  }

  const icons = { pending: "·", active: "…", done: "✓", error: "!" };

  const stepsHtml = steps.map((s) => {
    const st = stepState(s.key);
    const detail = s.detail[phase] || "";
    return `
      <li class="ws-step ${st}">
        <div class="ws-step-icon">${icons[st]}</div>
        <div class="ws-step-body">
          <p class="ws-step-title">${s.title}</p>
          ${detail ? `<p class="ws-step-detail">${detail}</p>` : ""}
        </div>
      </li>`;
  }).join("");

  let extraHtml = "";
  if (phase === "ready" && opts.summary) {
    const s = opts.summary;
    extraHtml = `
      <div class="ws-summary">
        <div><span class="ws-stat-val">${s.master_count ?? "—"}</span><span class="ws-stat-lbl">主表公司</span></div>
        <div><span class="ws-stat-val">${s.review_count ?? "—"}</span><span class="ws-stat-lbl">待確認</span></div>
        <div><span class="ws-stat-val">${s.candidate_count ?? "—"}</span><span class="ws-stat-lbl">新增候選</span></div>
      </div>
      <button class="ws-goto-btn" id="wsGotoBtn">前往總覽 →</button>`;
  }
  if (phase === "error" && opts.error) {
    extraHtml = `<div class="ws-error-msg">${opts.error}</div>`;
  }

  el.innerHTML = `
    <p class="workspace-eyebrow">工作進度</p>
    <ul class="ws-steps">${stepsHtml}</ul>
    ${extraHtml}`;

  // 綁 goto 按鈕
  document.getElementById("wsGotoBtn")?.addEventListener("click", () => setView("overview"));
}

async function createTaskFromUpload(onStatus) {
  const formData = new FormData();
  formData.append("task_name", elements.taskNameInput.value.trim() || "未命名任務");
  formData.append("chart1", state.chart1File);
  formData.append("chart2", state.chart2File);

  renderWorkspace("uploading");
  if (onStatus) onStatus("上傳中…");
  const initTask = await apiPost("/api/tasks/analyze", formData, true);

  renderWorkspace("processing", { msg: "AI 辨識中，請稍候…" });
  if (onStatus) onStatus("AI 辨識中，請稍候…");
  const elapsed = (tick) => {
    const secs = tick * 4;
    const msg = `Qwen-VL 辨識中… 已等候 ${secs} 秒`;
    renderWorkspace("processing", { msg });
    if (onStatus) onStatus(msg);
  };

  try {
    const task = await pollTask(initTask.id, elapsed);
    hydrateTask(task);
    renderWorkspace("ready", { summary: task.summary });
    // 不立刻切頁，讓使用者看到摘要後自行點「前往總覽」
  } catch (err) {
    renderWorkspace("error", { error: err.message });
    throw err;
  }
}

async function createDemoTask() {
  const task = await apiGet("/api/demo-task");
  hydrateTask(task);
  setView("overview");
}

function exportWorkbook() {
  if (!window.XLSX) {
    alert("目前無法載入匯出元件，請稍後再試。");
    return;
  }

  const summaryRows = [
    ["任務名稱", state.taskName || "未命名任務"],
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

// ══════════════════════════════════════════════════════════════
// 股權架構圖
// ══════════════════════════════════════════════════════════════

const LEVEL_COLORS = ["#1e3a5f", "#1d4ed8", "#0891b2", "#0d9488", "#059669", "#d97706"];
const LEVEL_NAMES  = ["頂層主體", "一級子公司", "二級子公司", "三級子公司", "四級子公司", "五級以上"];
let _chart = null;

// 節點尺寸
const NODE_W = 220;
const NODE_H = 110;

function wrapName(name, maxLen = 12) {
  if (!name || name.length <= maxLen) return name || "";
  const lines = [];
  for (let i = 0; i < name.length; i += maxLen) lines.push(name.slice(i, i + maxLen));
  return lines.join("\n");
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
        formatter: r.actual_controller_share,
        fontSize: 12,
        fontWeight: "bold",
        color: "#1e293b",
        backgroundColor: "#ffffff",
        padding: [3, 7],
        borderRadius: 4,
        borderWidth: 1,
        borderColor: "#cbd5e1",
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
  const roots = buildTree(state.masterRows);

  // 更新圖例
  const levels = [...new Set(state.masterRows.map((r) => Number(r.chart1_level) || 0))].sort();
  elements.chartLegend.innerHTML = [
    ...levels.map((l) => {
      const color = LEVEL_COLORS[Math.min(l, LEVEL_COLORS.length - 1)];
      return `<span class="legend-item"><span class="legend-dot" style="background:${color}"></span>${LEVEL_NAMES[l] || `第${l}層`}</span>`;
    }),
    `<span class="legend-item"><span class="legend-dot legend-dot-uncertain"></span>待確認</span>`,
  ].join("");

  let rows = [];

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
    const attrs = [
      node.legal_representative ? `法代：${node.legal_representative}` : "",
      node.registered_capital   ? `資本：${formatCapital(node.registered_capital)}` : "",
      node.established_date     ? `成立：${node.established_date}` : "",
    ].filter(Boolean).join("｜");

    rows.push({ prefix, name, attrs, color, uncertain, isRoot, level });

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
          <span class="lt-name" style="color:${r.color}">${r.name}</span>
          ${r.uncertain ? '<span class="lt-warn">⚠ 待確認</span>' : ""}
          ${r.attrs ? `<span class="lt-attrs">${r.attrs}</span>` : ""}
        </div>
      `).join("")}
    </div>`;
}

// ── ECharts 視覺圖（≤20 家） ─────────────────────────────────
function renderEChart() {
  if (!window.echarts) { console.warn("ECharts 未載入"); return; }
  if (_chart) { _chart.dispose(); _chart = null; }

  const container = elements.chartContainer;
  _chart = echarts.init(container, null, { renderer: "canvas" });

  const treeData = buildEChartsTree(state.masterRows);
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

  const levels = [...new Set(state.masterRows.map((r) => Number(r.chart1_level) || 0))].sort();
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

  const total = state.masterRows.length;
  const useList = total > 20;

  elements.chartLayoutBadge.textContent = `${total} 家公司 · ${useList ? "條列樹狀" : "視覺圖"}`;

  // 切換容器樣式
  elements.chartContainer.classList.toggle("chart-container-list", useList);
  elements.chartContainer.classList.toggle("chart-container-echart", !useList);

  // PNG / HTML 匯出按鈕只在 ECharts 模式下有意義
  elements.exportPngBtn.style.display  = useList ? "none" : "";
  elements.exportHtmlBtn.style.display = useList ? "none" : "";

  if (useList) renderListTree();
  else renderEChart();
}

function exportPNG() {
  if (!_chart) return;
  const url = _chart.getDataURL({ type: "png", pixelRatio: 2, backgroundColor: "#f8fafc" });
  const a = document.createElement("a");
  a.href = url;
  a.download = `${state.taskName || "股權架構圖"}.png`;
  a.click();
}

function exportHTML() {
  const title = state.taskName || "股權架構圖";
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

function printChart() {
  window.print();
}

function bindEvents() {
  elements.navButtons.forEach((button) => {
    button.addEventListener("click", () => setView(button.dataset.view));
  });
  elements.chart1Input.addEventListener("change", (event) => {
    state.chart1File = event.target.files[0];
    setPreview(state.chart1File, elements.chart1Meta, elements.chart1Preview, document.getElementById("dz1"));
    enableStartIfReady();
  });
  elements.chart2Input.addEventListener("change", (event) => {
    state.chart2File = event.target.files[0];
    setPreview(state.chart2File, elements.chart2Meta, elements.chart2Preview, document.getElementById("dz2"));
    enableStartIfReady();
  });
  elements.taskNameInput.addEventListener("input", (event) => {
    state.taskName = event.target.value.trim();
  });
  elements.startAnalysisBtn.addEventListener("click", async () => {
    const originalText = elements.startAnalysisBtn.textContent;
    document.getElementById("uploadError")?.remove();
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
  elements.loadDemoBtn.addEventListener("click", async () => {
    try {
      await createDemoTask();
    } catch (error) {
      console.error(error);
      alert("目前無法載入示範任務。");
    }
  });
  elements.searchInput.addEventListener("input", renderResults);
  elements.statusFilter.addEventListener("change", renderResults);
  elements.exportBtn.addEventListener("click", exportWorkbook);
  elements.exportPngBtn.addEventListener("click", exportPNG);
  elements.exportHtmlBtn.addEventListener("click", exportHTML);
  elements.printChartBtn.addEventListener("click", printChart);
}

bindEvents();
updateTaskBadge();

// 頁面載入時靜默 ping 後端，提前喚醒 Railway（冷啟動可能需 10–30 秒）
fetch(API_BASE + "/api/health").catch(() => {});

