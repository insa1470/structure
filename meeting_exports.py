from __future__ import annotations

import html
import re
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from xml.sax.saxutils import escape as xml_escape


SLIDE_W = 13.333
SLIDE_H = 7.5
EMU = 914400


def generate_meeting_html(task: dict, output_path: str | Path) -> dict:
    rows = _normalized_rows(task.get("master_rows") or [])
    title = _title(task)
    overview_svg = _build_overview_svg(title, rows)
    hierarchy = _ordered_rows(rows)
    table_rows = "\n".join(_html_table_row(row) for row, _depth in hierarchy)
    hierarchy_rows = "\n".join(_html_hierarchy_row(row, depth) for row, depth in hierarchy)
    max_depth = max((depth for _row, depth in hierarchy), default=0)
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(
        f"""<!doctype html>
<html lang="zh-Hant">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>{html.escape(title)}｜股權架構會議頁</title>
<style>
@page {{ size: A4 landscape; margin: 10mm; }}
* {{ box-sizing: border-box; }}
body {{
  margin: 0;
  background: #eef3f8;
  color: #172033;
  font-family: "Microsoft JhengHei", "PingFang TC", "Noto Sans TC", Arial, sans-serif;
}}
.meeting-page {{ max-width: 1500px; margin: 0 auto; padding: 24px; }}
.topbar {{ display: flex; justify-content: space-between; gap: 16px; align-items: baseline; margin-bottom: 14px; }}
h1 {{ font-size: 28px; margin: 0; letter-spacing: 0; }}
.meta {{ color: #64748b; font-size: 13px; white-space: nowrap; }}
.toolbar {{
  position: sticky;
  top: 0;
  z-index: 10;
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: 10px;
  margin: 0 0 14px;
  padding: 12px;
  background: rgba(238, 243, 248, .92);
  backdrop-filter: blur(10px);
  border: 1px solid #d8e0ea;
  border-radius: 8px;
}}
.toolbar input, .toolbar select {{
  height: 36px;
  border: 1px solid #cbd5e1;
  border-radius: 8px;
  padding: 0 10px;
  background: #fff;
  color: #172033;
  font: inherit;
}}
.toolbar input {{ min-width: 260px; }}
.toolbar button {{
  height: 36px;
  border: 1px solid #cbd5e1;
  border-radius: 8px;
  padding: 0 12px;
  background: #fff;
  color: #172033;
  font-weight: 700;
  cursor: pointer;
}}
.toolbar button:hover {{ border-color: #64748b; }}
.panel {{ background: #fff; border: 1px solid #d8e0ea; border-radius: 6px; padding: 14px; }}
.panel h2 {{ margin: 0 0 10px; font-size: 15px; }}
.meeting-stage {{ position: relative; min-height: calc(100vh - 146px); }}
.chart-panel {{
  min-height: calc(100vh - 170px);
  overflow: auto;
  display: flex;
  flex-direction: column;
}}
.chart-shell {{
  flex: 1;
  min-height: 560px;
  display: grid;
  place-items: center;
  background: #f8fafc;
  border: 1px solid #e2e8f0;
  border-radius: 6px;
  padding: 10px;
}}
.chart-panel svg {{ width: min(100%, 1380px); height: auto; }}
.chart-panel svg [data-node-id].is-selected rect {{ stroke: #f59e0b; stroke-width: 4; }}
.detail-drawer {{
  position: fixed;
  top: 0;
  right: 0;
  z-index: 30;
  width: clamp(320px, 28vw, 420px);
  height: 100vh;
  display: flex;
  flex-direction: column;
  background: rgba(255, 255, 255, .98);
  border-left: 1px solid #cbd5e1;
  box-shadow: -18px 0 36px rgba(15, 23, 42, .16);
  transform: translateX(100%);
  transition: transform 180ms ease;
}}
body.drawer-open .detail-drawer {{ transform: translateX(0); }}
.drawer-head {{
  display: flex;
  justify-content: space-between;
  gap: 10px;
  align-items: center;
  padding: 16px;
  border-bottom: 1px solid #e2e8f0;
}}
.drawer-head h2 {{ margin: 0; font-size: 16px; }}
.drawer-close {{
  width: 34px;
  height: 34px;
  border: 1px solid #cbd5e1;
  border-radius: 8px;
  background: #fff;
  font-weight: 800;
  cursor: pointer;
}}
.drawer-body {{ overflow: auto; padding: 12px 10px 18px; }}
.hierarchy {{ font-size: 13px; line-height: 1.78; overflow-x: auto; padding-bottom: 8px; }}
.h-row {{
  display: flex;
  align-items: baseline;
  gap: 6px;
  min-width: max-content;
  padding: 3px 8px;
  border-radius: 4px;
  white-space: nowrap;
  cursor: pointer;
}}
.h-row:nth-child(even) {{ background: #f8fafc; }}
.h-row:hover {{ background: #eaf2ff; }}
.h-row.is-hidden {{ display: none; }}
.h-row.is-match {{ background: #fff7d6; outline: 1px solid #f5c542; }}
.h-row.is-selected {{ background: #ffedd5; outline: 2px solid #f59e0b; }}
.h-prefix {{ color: #94a3b8; font-family: "Courier New", Courier, monospace; white-space: pre; }}
.h-name {{ font-weight: 800; }}
.h-warn {{ color: #d97706; font-size: 12px; font-weight: 800; }}
.h-attrs {{ color: #475569; font-size: 12px; }}
.status-line {{ color: #64748b; font-size: 12px; margin: 8px 0 0; }}
table {{ width: 100%; border-collapse: collapse; font-size: 11px; }}
th, td {{ border-bottom: 1px solid #e8edf3; padding: 6px 8px; text-align: left; vertical-align: top; }}
th {{ background: #f8fafc; color: #475569; font-weight: 700; }}
.full-table {{ margin-top: 16px; max-height: 520px; overflow: auto; }}
tr.is-hidden {{ display: none; }}
tr.is-match {{ background: #fff7d6; }}
.print-note {{ margin-top: 10px; color: #64748b; font-size: 12px; }}
@media print {{
  body {{ background: #fff; }}
  .meeting-page {{ max-width: none; padding: 0; }}
  .toolbar, .detail-drawer {{ display: none; }}
  .panel {{ break-inside: avoid; }}
  .hierarchy {{ font-size: 11px; line-height: 1.6; overflow: visible; }}
  .h-row.is-hidden, tr.is-hidden {{ display: flex; }}
  tr.is-hidden {{ display: table-row; }}
  .full-table {{ break-before: page; }}
  .print-note {{ display: none; }}
}}
</style>
</head>
<body>
<main class="meeting-page">
  <div class="topbar">
    <h1>{html.escape(title)}｜股權架構會議頁</h1>
    <div class="meta">{len(rows)} 筆主表資料｜{datetime.now(timezone.utc).strftime("%Y-%m-%d")}</div>
  </div>
  <section class="toolbar" aria-label="會議頁工具列">
    <input id="searchBox" type="search" placeholder="搜尋公司、法代、持股或關鍵字" />
    <select id="levelFilter" aria-label="層級篩選">
      <option value="all">顯示全部層級</option>
      {''.join(f'<option value="{level}">只看到 L{level}</option>' for level in range(max_depth + 1))}
    </select>
    <button type="button" data-depth="2">只看前兩層</button>
    <button type="button" data-depth="all">展開全部</button>
    <button type="button" id="toggleDrawerBtn">清單</button>
    <button type="button" id="printBtn">列印 / 存 PDF</button>
    <span class="status-line" id="resultCount">{len(rows)} 筆資料</span>
  </section>
  <section class="meeting-stage" aria-label="股權架構簡報舞台">
    <article class="panel chart-panel">
      <h2>股權總覽</h2>
      <div class="chart-shell">{overview_svg}</div>
      <p class="print-note">主圖保留公司名稱與持股，條列層級可由右上「清單」開啟，不會壓縮股權圖。</p>
    </article>
  </section>
  <aside class="detail-drawer" id="detailDrawer" aria-label="條列層級清單">
    <div class="drawer-head">
      <h2>條列層級</h2>
      <button class="drawer-close" type="button" id="closeDrawerBtn" aria-label="關閉清單">×</button>
    </div>
    <div class="drawer-body">
      <div class="hierarchy" id="hierarchyList">{hierarchy_rows}</div>
      <p class="print-note">點選公司可高亮主圖節點；搜尋與層級篩選會同步更新本清單。</p>
    </div>
  </aside>
  <section class="panel full-table">
    <h2>結果主表摘要</h2>
    <table id="resultTable">
      <thead><tr><th>層級</th><th>公司名稱</th><th>父層</th><th>持股</th><th>狀態</th></tr></thead>
      <tbody>{table_rows}</tbody>
    </table>
  </section>
</main>
<script>
(function() {{
  const searchBox = document.getElementById("searchBox");
  const levelFilter = document.getElementById("levelFilter");
  const resultCount = document.getElementById("resultCount");
  const toggleDrawerBtn = document.getElementById("toggleDrawerBtn");
  const closeDrawerBtn = document.getElementById("closeDrawerBtn");
  const hierarchyRows = Array.from(document.querySelectorAll(".h-row"));
  const tableRows = Array.from(document.querySelectorAll("#resultTable tbody tr"));
  let maxDepth = null;
  let selectedNodeId = "";

  function normalize(value) {{
    return String(value || "")
      .trim()
      .toLowerCase()
      .replaceAll("華", "华")
      .replaceAll("錦", "锦")
      .replaceAll("運", "运")
      .replaceAll("動", "动")
      .replaceAll("貿", "贸")
      .replaceAll("實", "实")
      .replaceAll("業", "业")
      .replaceAll("強", "强")
      .replaceAll("達", "达")
      .replaceAll("絲", "丝")
      .replaceAll("裝", "装")
      .replaceAll("備", "备")
      .replaceAll("術", "术")
      .replaceAll("發", "发");
  }}

  function applyFilters() {{
    const q = normalize(searchBox.value);
    const level = levelFilter.value;
    let visible = 0;
    hierarchyRows.forEach((row) => {{
      const rowLevel = Number(row.dataset.level || 0);
      const text = normalize(row.dataset.search || row.textContent);
      const matchQuery = !q || text.includes(q);
      const matchLevel = level === "all" || String(rowLevel) === level;
      const matchDepth = maxDepth === null || rowLevel <= maxDepth;
      const show = matchQuery && matchLevel && matchDepth;
      row.classList.toggle("is-hidden", !show);
      row.classList.toggle("is-match", Boolean(q && matchQuery));
      if (show) visible += 1;
    }});
    tableRows.forEach((row) => {{
      const rowLevel = Number(row.dataset.level || 0);
      const text = normalize(row.dataset.search || row.textContent);
      const matchQuery = !q || text.includes(q);
      const matchLevel = level === "all" || String(rowLevel) === level;
      const matchDepth = maxDepth === null || rowLevel <= maxDepth;
      const show = matchQuery && matchLevel && matchDepth;
      row.classList.toggle("is-hidden", !show);
      row.classList.toggle("is-match", Boolean(q && matchQuery));
    }});
    resultCount.textContent = visible + " 筆顯示中";
  }}

  function setDrawer(open) {{
    document.body.classList.toggle("drawer-open", open);
    toggleDrawerBtn.textContent = open ? "收合清單" : "清單";
  }}

  function selectNode(nodeId) {{
    selectedNodeId = nodeId || "";
    hierarchyRows.forEach((row) => row.classList.toggle("is-selected", row.dataset.nodeId === selectedNodeId));
    document.querySelectorAll("svg [data-node-id]").forEach((node) => {{
      node.classList.toggle("is-selected", node.dataset.nodeId === selectedNodeId);
    }});
  }}

  searchBox.addEventListener("input", applyFilters);
  levelFilter.addEventListener("change", applyFilters);
  toggleDrawerBtn.addEventListener("click", () => setDrawer(!document.body.classList.contains("drawer-open")));
  closeDrawerBtn.addEventListener("click", () => setDrawer(false));
  hierarchyRows.forEach((row) => {{
    row.addEventListener("click", () => selectNode(row.dataset.nodeId));
  }});
  document.querySelectorAll("[data-depth]").forEach((button) => {{
    button.addEventListener("click", () => {{
      maxDepth = button.dataset.depth === "all" ? null : Number(button.dataset.depth);
      applyFilters();
    }});
  }});
  document.getElementById("printBtn").addEventListener("click", () => window.print());
  setDrawer(false);
  applyFilters();
}})();
</script>
</body>
</html>
""",
        encoding="utf-8",
    )
    return {"ok": True, "path": str(output), "row_count": len(rows)}


def generate_meeting_pptx(task: dict, output_path: str | Path) -> dict:
    rows = _normalized_rows(task.get("master_rows") or [])
    title = _title(task)
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    slides = [_ppt_overview_slide(title, rows), *_ppt_hierarchy_slides(title, rows)]
    _write_pptx(output, slides, title)
    return {"ok": True, "path": str(output), "slide_count": len(slides), "row_count": len(rows)}


def _normalized_rows(rows: list[dict]) -> list[dict]:
    normalized = []
    for idx, row in enumerate(rows):
        item = dict(row or {})
        item["node_id"] = str(item.get("node_id") or f"N{idx + 1:03d}")
        item["canonical_name"] = _clean(item.get("canonical_name") or item.get("chart1_name") or "未命名公司")
        item["chart1_parent"] = str(item.get("chart1_parent") or "")
        item["chart1_level"] = _safe_int(item.get("chart1_level"), 0)
        item["actual_controller_share"] = _clean(item.get("actual_controller_share") or "")
        item["node_status"] = _clean(item.get("node_status") or "")
        normalized.append(item)
    return normalized


def _title(task: dict) -> str:
    return _clean(task.get("name") or task.get("id") or "股權架構")


def _clean(value) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def _safe_int(value, default: int = 0) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def _children_by_parent(rows: list[dict]) -> dict[str, list[dict]]:
    children: dict[str, list[dict]] = {}
    for row in rows:
        children.setdefault(str(row.get("chart1_parent") or ""), []).append(row)
    return children


def _root(rows: list[dict]) -> dict | None:
    by_id = {row["node_id"]: row for row in rows}
    return next((row for row in rows if not row.get("chart1_parent") or row.get("chart1_parent") not in by_id), rows[0] if rows else None)


def _ordered_rows(rows: list[dict]) -> list[tuple[dict, int]]:
    if not rows:
        return []
    children = _children_by_parent(rows)
    root = _root(rows)
    seen: set[str] = set()
    ordered: list[tuple[dict, int]] = []

    def walk(row: dict, depth: int) -> None:
        if row["node_id"] in seen:
            return
        seen.add(row["node_id"])
        ordered.append((row, depth))
        for child in children.get(row["node_id"], []):
            walk(child, depth + 1)

    if root:
        walk(root, 0)
    for row in rows:
        if row["node_id"] not in seen:
            walk(row, _safe_int(row.get("chart1_level"), 0))
    return ordered


def _overview_model(rows: list[dict]) -> dict:
    children = _children_by_parent(rows)
    root = _root(rows)
    main = (children.get(root["node_id"], []) or [root])[0] if root else None
    level2 = children.get(main["node_id"], []) if main else []
    return {
        "root": root,
        "main": main,
        "visible": level2[:5],
        "overflow": level2[5:],
        "children": children,
    }


def _build_overview_svg(title: str, rows: list[dict]) -> str:
    model = _overview_model(rows)
    root, main, visible, overflow, children = model["root"], model["main"], model["visible"], model["overflow"], model["children"]
    w, h = 1220, 560
    parts = [
        f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {w} {h}" width="100%" role="img" aria-label="{html.escape(title)} 股權總覽">',
        "<style>.n rect{fill:#fff;stroke:#172033;stroke-width:1.4}.root rect{fill:#e9edf8}.g rect{fill:#eaf2fb;stroke:#2f5f98;stroke-width:1.3}.e{stroke:#6b7c93;stroke-width:1.2}.t{font:600 13px sans-serif;fill:#172033}.s{font:500 11px sans-serif;fill:#172033}.gt{font:700 14px sans-serif;fill:#172033}</style>",
        '<rect x="1" y="1" width="1218" height="558" fill="#f8fafc" stroke="#d6dee8"/>',
    ]
    if not root:
        parts.append('<text x="610" y="280" text-anchor="middle" class="t">沒有可顯示的公司資料</text></svg>')
        return "".join(parts)
    parts.append(_svg_node(root["node_id"], root["canonical_name"], 475, 28, 270, 48, root=True))
    if main and main is not root:
        parts.append(_svg_node(main["node_id"], main["canonical_name"], 475, 112, 270, 54))
        parts.append(_svg_line(610, 76, 610, 112))
    top_y = 252
    xs = [48, 276, 504, 732, 960]
    if visible:
        parts.append(_svg_line(610, 166, 610, 218))
        parts.append(_svg_line(xs[0] + 95, 218, xs[len(visible) - 1] + 95, 218))
    for idx, row in enumerate(visible):
        x = xs[idx]
        parts.append(_svg_line(x + 95, 218, x + 95, top_y))
        parts.append(_svg_node(row["node_id"], _node_label(row), x, top_y, 190, 58))
        kids = children.get(row["node_id"], [])
        if kids:
            parts.append(_svg_line(x + 95, top_y + 58, x + 95, 342))
            parts.append(_svg_group(f"清單柱｜{len(kids)} 家", kids, x - 8, 354, 206, 160, max_items=5))
    if overflow:
        parts.append(_svg_line(xs[-1] + 95, 218, xs[-1] + 95, 374))
        parts.append(_svg_group(f"其他一級子公司｜{len(overflow)} 家", overflow, 928, 384, 238, 150, max_items=5))
    parts.append("</svg>")
    return "".join(parts)


def _svg_node(node_id: str, text: str, x: int, y: int, w: int, h: int, root: bool = False) -> str:
    lines = _split_label(text, 14)
    cls = "n root" if root else "n"
    ty = y + h / 2 - (len(lines) - 1) * 7
    text_xml = "".join(f'<text x="{x + w / 2}" y="{ty + i * 16}" text-anchor="middle" dominant-baseline="middle" class="t">{html.escape(line)}</text>' for i, line in enumerate(lines))
    return f'<g class="{cls}" data-node-id="{html.escape(node_id)}"><rect x="{x}" y="{y}" width="{w}" height="{h}" rx="2"/>{text_xml}</g>'


def _svg_group(title: str, rows: list[dict], x: int, y: int, w: int, h: int, max_items: int = 5) -> str:
    items = rows[:max_items]
    lines = [f'<g class="g"><rect x="{x}" y="{y}" width="{w}" height="{h}" rx="2"/><text x="{x + 12}" y="{y + 24}" class="gt">{html.escape(title)}</text>']
    for idx, row in enumerate(items):
        label = f"{idx + 1}. {_truncate(row['canonical_name'], 13)}"
        if row.get("actual_controller_share"):
            label += f" {row['actual_controller_share']}"
        lines.append(f'<text x="{x + 14}" y="{y + 50 + idx * 18}" class="s">{html.escape(label)}</text>')
    if len(rows) > len(items):
        lines.append(f'<text x="{x + 14}" y="{y + 50 + len(items) * 18}" class="s">另有 {len(rows) - len(items)} 家未列示</text>')
    lines.append("</g>")
    return "".join(lines)


def _svg_line(x1, y1, x2, y2) -> str:
    return f'<line class="e" x1="{x1}" y1="{y1}" x2="{x2}" y2="{y2}"/>'


def _html_hierarchy_row(row: dict, depth: int) -> str:
    prefix = "&nbsp;" * (depth * 4) + ("└── " if depth else "")
    attrs = html.escape(_row_attrs(row))
    color = "#" + _ppt_level_text_color(depth)
    uncertain = row.get("node_status") != "enriched"
    search = html.escape(" ".join([row["canonical_name"], _row_attrs(row), row.get("node_status") or ""]), quote=True)
    warn_html = '<span class="h-warn">⚠ 待確認</span>' if uncertain else ""
    attrs_html = f'<span class="h-attrs">{attrs}</span>' if attrs else ""
    return (
        f'<div class="h-row" data-level="{depth}" data-node-id="{html.escape(row["node_id"], quote=True)}" data-search="{search}">'
        f'<span class="h-prefix">{prefix}</span>'
        f'<span class="h-name" style="color:{color}">{html.escape(row["canonical_name"])}</span>'
        f"{warn_html}"
        f"{attrs_html}"
        "</div>"
    )


def _html_table_row(row: dict) -> str:
    level = _safe_int(row.get("chart1_level"), 0)
    search = html.escape(" ".join([row["canonical_name"], _row_attrs(row), row.get("node_status") or ""]), quote=True)
    return (
        f'<tr data-level="{level}" data-node-id="{html.escape(row["node_id"], quote=True)}" data-search="{search}">'
        f"<td>L{level}</td>"
        f"<td>{html.escape(row['canonical_name'])}</td>"
        f"<td>{html.escape(row.get('chart1_parent') or '')}</td>"
        f"<td>{html.escape(row.get('actual_controller_share') or '')}</td>"
        f"<td>{html.escape(row.get('node_status') or '')}</td>"
        "</tr>"
    )


def _node_label(row: dict) -> str:
    share = row.get("actual_controller_share") or ""
    return f"{row['canonical_name']}\n持股：{share}" if share else row["canonical_name"]


def _split_label(text: str, width: int) -> list[str]:
    chunks = []
    for raw in str(text).split("\n"):
        raw = raw.strip()
        if len(raw) <= width:
            chunks.append(raw)
        else:
            chunks.append(raw[:width])
            chunks.append(raw[width : width * 2] + ("…" if len(raw) > width * 2 else ""))
    return chunks[:3]


def _truncate(text: str, width: int) -> str:
    return text if len(text) <= width else text[: width - 1] + "…"


def _ppt_overview_slide(title: str, rows: list[dict]) -> str:
    model = _overview_model(rows)
    root, main, visible, overflow, children = model["root"], model["main"], model["visible"], model["overflow"], model["children"]
    shapes = [_ppt_text(title + "｜股權架構會議圖", 0.35, 0.24, 8.8, 0.35, 18, bold=True), _ppt_text("可編輯 PPT：公司框、線條、清單柱與文字皆為原生物件。", 0.36, 0.62, 10.8, 0.22, 8.5, color="667085"), _ppt_rect(0.34, 0.92, 12.65, 5.92, fill="F7FAFC", line="D6DEE8")]
    if not root:
        shapes.append(_ppt_text("沒有可顯示的公司資料", 5.0, 3.4, 3.0, 0.3, 12))
        return "".join(shapes)
    shapes.append(_ppt_node(root["canonical_name"], 5.15, 1.05, 2.75, 0.42, fill="E9EDF8", bold=True))
    if main and main is not root:
        shapes.append(_ppt_node(main["canonical_name"], 5.15, 1.78, 2.75, 0.48, bold=True))
        shapes.append(_ppt_line(6.525, 1.47, 6.525, 1.78))
    xs = [0.65, 3.05, 5.45, 7.85, 10.25]
    box_w = 2.05
    top_y = 3.0
    bus_y = 2.63
    if visible:
        shapes.append(_ppt_line(6.525, 2.26, 6.525, bus_y))
        shapes.append(_ppt_line(xs[0] + box_w / 2, bus_y, xs[len(visible) - 1] + box_w / 2, bus_y))
    for idx, row in enumerate(visible):
        x = xs[idx]
        shapes.append(_ppt_line(x + box_w / 2, bus_y, x + box_w / 2, top_y))
        shapes.append(_ppt_node(_node_label(row), x, top_y, box_w, 0.56, font_size=6.8))
        kids = children.get(row["node_id"], [])
        if kids:
            shapes.append(_ppt_line(x + box_w / 2, top_y + 0.56, x + box_w / 2, 3.83))
            shapes.append(_ppt_group_box(f"清單柱｜{len(kids)} 家", kids, x - 0.1, 3.88, 2.25, 1.36))
    if overflow:
        shapes.append(_ppt_line(11.32, bus_y, 11.32, 5.22))
        shapes.append(_ppt_group_box(f"其他一級子公司｜{len(overflow)} 家", overflow, 9.95, 5.28, 2.75, 1.08, max_items=5))
    shapes.append(_ppt_text("用戶可編輯：拖曳公司框、改名稱／持股、刪除或新增線條、調整清單柱內容。", 0.44, 6.96, 12.2, 0.2, 8, color="667085"))
    return "".join(shapes)


def _ppt_hierarchy_slides(title: str, rows: list[dict]) -> list[str]:
    tree_rows = _tree_list_rows(rows)
    if not tree_rows:
        return [_ppt_empty_hierarchy_slide(title)]
    per_slide = 26
    pages = [tree_rows[idx : idx + per_slide] for idx in range(0, len(tree_rows), per_slide)]
    return [_ppt_tree_list_slide(title, page, page_idx + 1, len(pages), len(tree_rows)) for page_idx, page in enumerate(pages)]


def _tree_list_rows(rows: list[dict]) -> list[dict]:
    children = _children_by_parent(rows)
    roots = [row for row, _depth in _ordered_rows(rows) if not row.get("chart1_parent")]
    if not roots and rows:
        root = _root(rows)
        roots = [root] if root else []
    output: list[dict] = []
    seen: set[str] = set()

    def walk(node: dict, prefix_lines: list[str], is_last: bool) -> None:
        if node["node_id"] in seen:
            return
        seen.add(node["node_id"])
        level = _safe_int(node.get("chart1_level"), 0)
        is_root = level == 0 and not prefix_lines
        connector = "" if is_root else ("└── " if is_last else "├── ")
        prefix = "".join(prefix_lines) + connector
        output.append(
            {
                "prefix": prefix,
                "name": node["canonical_name"],
                "attrs": _row_attrs(node),
                "level": level,
                "color": _ppt_level_text_color(level),
                "uncertain": node.get("node_status") != "enriched",
                "is_root": is_root,
            }
        )
        kids = children.get(node["node_id"], [])
        if kids:
            child_base = "" if is_root else ("    " if is_last else "│   ")
            for idx, child in enumerate(kids):
                walk(child, [*prefix_lines, child_base], idx == len(kids) - 1)

    for idx, root in enumerate(roots):
        if root:
            walk(root, [], idx == len(roots) - 1)
    for row in rows:
        if row["node_id"] not in seen:
            walk(row, [], True)
    return output


def _ppt_empty_hierarchy_slide(title: str) -> str:
    return "".join(
        [
            _ppt_text(title + "｜父子關係索引", 0.35, 0.24, 8.8, 0.35, 18, bold=True),
            _ppt_rect(0.34, 0.92, 12.65, 5.92, fill="F7FAFC", line="D6DEE8"),
            _ppt_text("沒有可顯示的父子關係資料。", 5.0, 3.4, 3.3, 0.3, 12, align="ctr"),
        ]
    )


def _ppt_tree_list_slide(title: str, items: list[dict], page: int, total_pages: int, total_count: int) -> str:
    page_label = f"（{page}/{total_pages}）" if total_pages > 1 else ""
    shapes = [
        _ppt_text(title + f"｜條列層級{page_label}", 0.35, 0.24, 8.8, 0.35, 18, bold=True),
        _ppt_text("單頁會議版：沿用網頁條列層級，一列一個文字框，方便直接編輯。", 0.36, 0.62, 10.8, 0.22, 8.5, color="667085"),
        _ppt_text(f"{total_count} 筆主體", 11.0, 0.33, 1.8, 0.2, 8.5, color="667085", align="r"),
        _ppt_rect(0.34, 0.92, 12.65, 5.92, fill="F7FAFC", line="D6DEE8"),
    ]
    shapes.extend(_ppt_tree_list_rows(items, 0.58, 1.12, 12.1))
    shapes.append(_ppt_text("每一列都是可編輯文字框；若要調整層級關係，建議回到結果主表修改後重新輸出。", 0.44, 6.96, 12.2, 0.2, 8, color="667085"))
    return "".join(shapes)


def _ppt_tree_list_rows(items: list[dict], x: float, y: float, w: float) -> list[str]:
    shapes: list[str] = []
    row_h = min(0.255, 5.72 / max(len(items), 1))
    for idx, row in enumerate(items):
        yy = y + idx * row_h
        fill = "FFFFFF" if idx % 2 == 0 else "F8FAFC"
        shapes.append(_ppt_rect(x - 0.08, yy - 0.018, w, row_h * 0.88, fill=fill, line=fill))
        shapes.append(_ppt_rect(x - 0.08, yy - 0.018, 0.045, row_h * 0.88, fill=row["color"], line=row["color"]))
        line = _ppt_tree_row_line(row)
        size = 6.45 if not row["is_root"] else 7.05
        shapes.append(_ppt_text(line, x, yy + 0.035, w - 0.16, 0.13, size, bold=row["is_root"], color="172033"))
    return shapes


def _ppt_tree_row_line(row: dict) -> str:
    warn = " ⚠待確認" if row["uncertain"] else ""
    prefix = row["prefix"]
    name = _truncate(row["name"], 18)
    attrs = _truncate(row["attrs"], 82)
    spacer = "  " if attrs else ""
    return f"{prefix}{name}{warn}{spacer}{attrs}"


def _row_attrs(row: dict) -> str:
    attrs = [
        f"法代：{row['legal_representative']}" if row.get("legal_representative") else "",
        f"資本：{row['registered_capital']}" if row.get("registered_capital") else "",
        f"成立：{row['established_date']}" if row.get("established_date") else "",
        f"持股：{row['actual_controller_share']}" if row.get("actual_controller_share") else "",
        f"定位：{row['role_label']}" if row.get("role_label") else "",
        f"備註：{row['chart_note']}" if row.get("chart_note") else "",
    ]
    return "｜".join(item for item in attrs if item)


def _ppt_level_text_color(depth: int) -> str:
    palette = ["1E3A5F", "1D4ED8", "0891B2", "0D9488", "059669", "D97706"]
    return palette[min(max(depth, 0), len(palette) - 1)]


def _ppt_group_box(title: str, rows: list[dict], x: float, y: float, w: float, h: float, max_items: int = 6) -> str:
    items = rows[:max_items]
    body = "\n".join(f"{idx + 1}. {_truncate(row['canonical_name'], 14)}{('  ' + row['actual_controller_share']) if row.get('actual_controller_share') else ''}" for idx, row in enumerate(items))
    if len(rows) > len(items):
        body += f"\n另有 {len(rows) - len(items)} 家未列示"
    return _ppt_rect(x, y, w, h, fill="EAF2FB", line="2F5F98") + _ppt_text(title, x + 0.10, y + 0.08, w - 0.2, 0.2, 8.3, bold=True) + _ppt_text(body, x + 0.12, y + 0.40, w - 0.24, h - 0.46, 6.4)


def _ppt_node(text: str, x: float, y: float, w: float, h: float, fill: str = "FFFFFF", bold: bool = False, font_size: float = 8.2) -> str:
    return _ppt_rect(x, y, w, h, fill=fill, line="172033") + _ppt_text(text, x + 0.05, y + 0.05, w - 0.1, h - 0.1, font_size, bold=bold, align="ctr", valign="mid")


def _ppt_rect(x: float, y: float, w: float, h: float, fill: str, line: str = "172033") -> str:
    return _ppt_shape("rect", x, y, w, h, fill=fill, line=line)


def _ppt_line(x1: float, y1: float, x2: float, y2: float, color: str = "6B7C93", width: float = 1.0) -> str:
    x, y = min(x1, x2), min(y1, y2)
    w, h = max(abs(x2 - x1), 0.006), max(abs(y2 - y1), 0.006)
    return _ppt_shape("rect", x, y, w, h, fill=color, line=color, line_width=width)


def _ppt_text(text: str, x: float, y: float, w: float, h: float, size: float, bold: bool = False, color: str = "172033", align: str = "l", valign: str = "top") -> str:
    paragraphs = []
    for line in str(text).split("\n"):
        paragraphs.append(
            f'<a:p><a:pPr algn="{align}"/><a:r><a:rPr lang="zh-TW" sz="{int(size * 100)}" b="{1 if bold else 0}"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:latin typeface="Microsoft JhengHei"/><a:ea typeface="Microsoft JhengHei"/></a:rPr><a:t>{xml_escape(line)}</a:t></a:r></a:p>'
        )
    anchor = "ctr" if valign == "mid" else "t"
    return _ppt_shape("rect", x, y, w, h, fill=None, line=None, text="".join(paragraphs), anchor=anchor)


def _ppt_shape(prst: str, x: float, y: float, w: float, h: float, fill: str | None, line: str | None, text: str = "", anchor: str = "t", line_width: float = 1.0) -> str:
    spid = _ppt_shape.next_id
    _ppt_shape.next_id += 1
    fill_xml = '<a:noFill/>' if fill is None else f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
    line_xml = '<a:ln><a:noFill/></a:ln>' if line is None else f'<a:ln w="{int(line_width * 12700)}"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill></a:ln>'
    body = f'<p:txBody><a:bodyPr wrap="square" anchor="{anchor}"/><a:lstStyle/>{text or "<a:p/>"}</p:txBody>'
    return f'<p:sp><p:nvSpPr><p:cNvPr id="{spid}" name="Shape {spid}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="{_emu(x)}" y="{_emu(y)}"/><a:ext cx="{_emu(w)}" cy="{_emu(h)}"/></a:xfrm><a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom>{fill_xml}{line_xml}</p:spPr>{body}</p:sp>'


_ppt_shape.next_id = 2


def _emu(value: float) -> int:
    return int(round(value * EMU))


def _write_pptx(path: Path, slide_xml_list: list[str], title: str) -> None:
    _ppt_shape.next_id = 2
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", _content_types(len(slide_xml_list)))
        z.writestr("_rels/.rels", _root_rels())
        z.writestr("docProps/core.xml", _core_xml(title))
        z.writestr("docProps/app.xml", _app_xml(len(slide_xml_list)))
        z.writestr("ppt/presentation.xml", _presentation_xml(len(slide_xml_list)))
        z.writestr("ppt/_rels/presentation.xml.rels", _presentation_rels(len(slide_xml_list)))
        z.writestr("ppt/theme/theme1.xml", _theme_xml())
        z.writestr("ppt/slideMasters/slideMaster1.xml", _slide_master_xml())
        z.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", _slide_master_rels())
        z.writestr("ppt/slideLayouts/slideLayout1.xml", _slide_layout_xml())
        z.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels", _slide_layout_rels())
        for idx, content in enumerate(slide_xml_list, start=1):
            z.writestr(f"ppt/slides/slide{idx}.xml", _slide_xml(content))
            z.writestr(f"ppt/slides/_rels/slide{idx}.xml.rels", _slide_rels())


def _content_types(slides: int) -> str:
    slide_overrides = "".join(f'<Override PartName="/ppt/slides/slide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>' for i in range(1, slides + 1))
    return f'<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>{slide_overrides}</Types>'


def _root_rels() -> str:
    return '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>'


def _core_xml(title: str) -> str:
    now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    return f'<?xml version="1.0" encoding="UTF-8"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>{xml_escape(title)}</dc:title><dc:creator>Structure</dc:creator><cp:lastModifiedBy>Structure</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified></cp:coreProperties>'


def _app_xml(slides: int) -> str:
    return f'<?xml version="1.0" encoding="UTF-8"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Structure</Application><PresentationFormat>Wide</PresentationFormat><Slides>{slides}</Slides></Properties>'


def _presentation_xml(slides: int) -> str:
    ids = "".join(f'<p:sldId id="{255 + i}" r:id="rId{i}"/>' for i in range(1, slides + 1))
    return f'<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId{slides + 1}"/></p:sldMasterIdLst><p:sldIdLst>{ids}</p:sldIdLst><p:sldSz cx="{_emu(SLIDE_W)}" cy="{_emu(SLIDE_H)}" type="wide"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>'


def _presentation_rels(slides: int) -> str:
    rels = [f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i}.xml"/>' for i in range(1, slides + 1)]
    rels.append(f'<Relationship Id="rId{slides + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>')
    rels.append(f'<Relationship Id="rId{slides + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>')
    return f'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{"".join(rels)}</Relationships>'


def _slide_xml(content: str) -> str:
    return f'<?xml version="1.0" encoding="UTF-8"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>{content}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'


def _slide_rels() -> str:
    return '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>'


def _slide_master_xml() -> str:
    return '<?xml version="1.0" encoding="UTF-8"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>'


def _slide_master_rels() -> str:
    return '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>'


def _slide_layout_xml() -> str:
    return '<?xml version="1.0" encoding="UTF-8"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank"><p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>'


def _slide_layout_rels() -> str:
    return '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>'


def _theme_xml() -> str:
    return '<?xml version="1.0" encoding="UTF-8"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Structure"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:srgbClr val="172033"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F2937"/></a:dk2><a:lt2><a:srgbClr val="F7FAFC"/></a:lt2><a:accent1><a:srgbClr val="2F5F98"/></a:accent1><a:accent2><a:srgbClr val="64748B"/></a:accent2><a:accent3><a:srgbClr val="EAF2FB"/></a:accent3><a:accent4><a:srgbClr val="D6DEE8"/></a:accent4><a:accent5><a:srgbClr val="94A3B8"/></a:accent5><a:accent6><a:srgbClr val="CBD5E1"/></a:accent6><a:hlink><a:srgbClr val="2F5F98"/></a:hlink><a:folHlink><a:srgbClr val="2F5F98"/></a:folHlink></a:clrScheme><a:fontScheme name="Structure"><a:majorFont><a:latin typeface="Microsoft JhengHei"/><a:ea typeface="Microsoft JhengHei"/></a:majorFont><a:minorFont><a:latin typeface="Microsoft JhengHei"/><a:ea typeface="Microsoft JhengHei"/></a:minorFont></a:fontScheme><a:fmtScheme name="Structure"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>'
