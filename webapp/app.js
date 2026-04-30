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
  if (!response.ok) throw new Error(`POST ${url} failed`);
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

function setPreview(file, metaEl, imgEl) {
  if (!file) return;
  metaEl.textContent = `${file.name} · ${(file.size / 1024 / 1024).toFixed(2)} MB`;
  const objectUrl = URL.createObjectURL(file);
  imgEl.src = objectUrl;
}

function makeMetric(label, value, theme) {
  return `
    <article class="metric-card ${theme}">
      <p class="metric-label">${label}</p>
      <p class="metric-value">${value}</p>
    </article>
  `;
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
  // 把開頭的數字部分加千分位，例如 28500萬元 → 28,500萬元
  return str.replace(/^(\d+)/, (_, n) => parseInt(n, 10).toLocaleString("en-US"));
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
  const query = elements.searchInput.value.trim();
  const filter = elements.statusFilter.value;

  const filteredRows = state.masterRows.filter((row) => {
    const inFilter = filter === "all" || row.node_status === filter;
    const haystack = [row.canonical_name, row.chart1_name, row.chart1_parent_name, row.legal_representative]
      .join(" ").toLowerCase();
    const inQuery = !query || haystack.includes(query.toLowerCase());
    return inFilter && inQuery;
  });

  // 搜尋時用篩選後的平面清單，否則用樹狀
  const rows = query || filter !== "all"
    ? filteredRows.map((r) => ({ ...r, _depth: 0 }))
    : flattenTree(buildTree(state.masterRows));

  elements.resultTableTitle.textContent = `${filteredRows.length} 家公司`;

  elements.resultTableBody.innerHTML = "";
  rows.forEach((row) => {
    const depth = row._depth || 0;
    const statusClass =
      row.node_status === "enriched" ? "status-enriched"
      : row.node_status === "review_match" ? "status-review"
      : "status-slate";

    const tr = document.createElement("tr");
    tr.className = statusClass;

    // 公司名稱（含縮排與層級線，可點擊編輯）
    const nameTd = document.createElement("td");
    nameTd.style.paddingLeft = `${8 + depth * 22}px`;
    nameTd.className = "tree-name-cell";
    const prefix = depth > 0 ? `<span class="tree-branch">${depth > 1 ? "　".repeat(depth - 1) + "└─ " : "└─ "}</span>` : "";
    const levelBadge = row.subsidiary_level_label
      ? `<span class="level-badge">${row.subsidiary_level_label}</span>` : "";
    nameTd.innerHTML = `${prefix}<span class="company-name editable-name" title="點擊編輯名稱">${row.canonical_name || row.chart1_name}</span>${levelBadge}`;

    // 名稱行內編輯
    nameTd.addEventListener("click", (e) => {
      if (e.target.classList.contains("tree-branch") || e.target.classList.contains("level-badge")) return;
      if (nameTd.querySelector("input.cell-input")) return;
      const nameSpan = nameTd.querySelector(".company-name");
      if (!nameSpan) return;
      const curName = row.canonical_name || row.chart1_name || "";
      nameSpan.innerHTML = `<input class="cell-input" value="${curName.replace(/"/g, "&quot;")}">`;
      const input = nameSpan.querySelector("input");
      input.focus(); input.select();
      const save = async () => {
        const newVal = input.value.trim();
        nameSpan.textContent = newVal || curName;
        if (newVal && newVal !== curName) {
          try {
            const result = await apiPost(`/api/tasks/${state.taskId}/update-row`, {
              node_id: row.node_id, canonical_name: newVal,
            });
            state.masterRows = result.master_rows || state.masterRows;
            row.canonical_name = newVal;
          } catch (e) { console.error(e); nameSpan.textContent = curName; }
        }
      };
      input.addEventListener("blur", save);
      input.addEventListener("keydown", (e) => {
        if (e.key === "Enter") { e.preventDefault(); input.blur(); }
        if (e.key === "Escape") { nameSpan.textContent = curName; }
      });
    });

    tr.appendChild(nameTd);

    // 可編輯欄位
    const editableFields = [
      { key: "legal_representative",    display: row.legal_representative },
      { key: "registered_capital",      display: formatCapital(row.registered_capital) },
      { key: "established_date",        display: row.established_date },
      { key: "actual_controller_share", display: row.actual_controller_share },
      { key: "company_status",          display: row.company_status },
    ];

    editableFields.forEach(({ key, display }) => {
      const td = document.createElement("td");
      td.textContent = display || "—";
      td.className = "editable-cell";
      makeEditable(td, row, key, display);
      tr.appendChild(td);
    });

    // 狀態（不可編輯）
    const statusTd = document.createElement("td");
    statusTd.textContent = statusText(row);
    tr.appendChild(statusTd);

    // 刪除按鈕
    const delTd = document.createElement("td");
    delTd.className = "del-td";
    const delBtn = document.createElement("button");
    delBtn.className = "row-del-btn";
    delBtn.title = "從主表移除此公司";
    delBtn.textContent = "✕";
    delBtn.addEventListener("click", async () => {
      if (!confirm(`確定要從主表移除「${row.canonical_name || row.chart1_name}」嗎？`)) return;
      try {
        const result = await apiPost(`/api/tasks/${state.taskId}/delete-row`, { node_id: row.node_id });
        state.masterRows = result.master_rows || state.masterRows;
        renderResults();
      } catch (e) { console.error("刪除失敗", e); }
    });
    delTd.appendChild(delBtn);
    tr.appendChild(delTd);

    elements.resultTableBody.appendChild(tr);
  });
}

async function createTaskFromUpload() {
  const formData = new FormData();
  formData.append("task_name", elements.taskNameInput.value.trim() || "未命名任務");
  formData.append("chart1", state.chart1File);
  formData.append("chart2", state.chart2File);
  const task = await apiPost("/api/tasks/analyze", formData, true);
  hydrateTask(task);
  setView("overview");
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

const LEVEL_COLORS = ["#1e3a5f", "#1d4ed8", "#4338ca", "#7c3aed", "#9333ea", "#c026d3"];
const LEVEL_NAMES  = ["頂層主體", "一級子公司", "二級子公司", "三級子公司", "四級子公司", "五級以上"];
let _chart = null;

function wrapName(name, maxLen = 10) {
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

    const nameLine = wrapName(r.canonical_name || r.chart1_name || "—");
    const labelParts = [`{name|${uncertain ? "⚠ " : ""}${nameLine}}`];
    if (r.legal_representative) labelParts.push(`{info|法代：${r.legal_representative}}`);
    if (r.registered_capital)   labelParts.push(`{info|資本：${formatCapital(r.registered_capital)}}`);
    if (r.established_date)     labelParts.push(`{info|成立：${r.established_date}}`);

    return {
      name: id,
      _row: r,
      label: { formatter: labelParts.join("\n") },
      itemStyle: {
        color,
        borderColor:  uncertain ? "#f59e0b" : color,
        borderWidth:  uncertain ? 2.5 : 0,
        borderType:   uncertain ? "dashed" : "solid",
        shadowColor:  "rgba(0,0,0,0.15)",
        shadowBlur:   6,
      },
      // 持股比例顯示在連線上
      edgeLabel: r.actual_controller_share
        ? { show: true, formatter: r.actual_controller_share, fontSize: 11, color: "#475569", backgroundColor: "#f1f5f9", padding: [2, 4], borderRadius: 3 }
        : undefined,
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
      top: "4%", bottom: "4%", left: "8%", right: "8%",
      symbol: "rect", symbolSize: [185, 90],
      edgeShape: "polyline", layout: "orthogonal",
      roam: true, initialTreeDepth: -1,
      label: {
        show: true, position: "inside",
        verticalAlign: "middle", align: "center",
        rich: {
          name: { fontSize: 12, fontWeight: "bold", color: "#fff", lineHeight: 20 },
          info: { fontSize: 10, color: "rgba(255,255,255,0.88)", lineHeight: 16 },
        },
      },
      leaves: { label: { position: "inside", verticalAlign: "middle", align: "center" } },
      lineStyle: { color: "#94a3b8", width: 1.5, curveness: 0 },
      emphasis: { focus: "descendant" },
      animationDurationUpdate: 600,
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
    setPreview(state.chart1File, elements.chart1Meta, elements.chart1Preview);
    enableStartIfReady();
  });
  elements.chart2Input.addEventListener("change", (event) => {
    state.chart2File = event.target.files[0];
    setPreview(state.chart2File, elements.chart2Meta, elements.chart2Preview);
    enableStartIfReady();
  });
  elements.taskNameInput.addEventListener("input", (event) => {
    state.taskName = event.target.value.trim();
  });
  elements.startAnalysisBtn.addEventListener("click", async () => {
    try {
      state.loading = true;
      enableStartIfReady();
      await createTaskFromUpload();
    } catch (error) {
      console.error(error);
      alert("目前分析服務無法完成任務建立。");
    } finally {
      state.loading = false;
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
