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
};

const pageTitles = {
  upload: "建立新任務",
  overview: "總覽",
  review: "待確認",
  candidates: "圖二新增候選",
  results: "結果主表",
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

function renderResults() {
  const query = elements.searchInput.value.trim();
  const filter = elements.statusFilter.value;
  const visibleRows = state.masterRows.filter((row) => {
    const inFilter = filter === "all" ? true : row.node_status === filter;
    const haystack = [row.chart1_name, row.chart1_parent_name, row.legal_representative, row.matched_chart2_name]
      .join(" ")
      .toLowerCase();
    const inQuery = query ? haystack.includes(query.toLowerCase()) : true;
    return inFilter && inQuery;
  });

  elements.resultTableTitle.textContent = `${visibleRows.length} 家公司`;
  elements.resultTableBody.innerHTML = visibleRows
    .map((row) => {
      const reviewDecision = state.reviewDecisions[row.node_id] || state.reviewDecisions[row.chart1_name] || {};
      const statusClass =
        row.node_status === "enriched"
          ? "status-enriched"
          : row.node_status === "review_match"
            ? "status-review"
            : "status-slate";
      return `
        <tr class="${statusClass}">
          <td>${reviewDecision.corrected_name || row.canonical_name || row.chart1_name}</td>
          <td>${reviewDecision.corrected_parent || row.chart1_parent_name || "—"}</td>
          <td>${reviewDecision.corrected_level || row.chart1_level || "—"}</td>
          <td>${row.subsidiary_level_label || "—"}</td>
          <td>${row.legal_representative || "—"}</td>
          <td>${row.established_date || "—"}</td>
          <td>${row.registered_capital || "—"}</td>
          <td>${row.actual_controller_share || "—"}</td>
          <td>${row.company_status || "—"}</td>
          <td>${statusText(row)}</td>
          <td>${reviewDecision.decision || "—"}</td>
          <td class="table-note">${reviewDecision.note || row.review_note || "—"}</td>
        </tr>
      `;
    })
    .join("");
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
}

bindEvents();
updateTaskBadge();
